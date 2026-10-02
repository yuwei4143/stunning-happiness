"""極簡 Excel 公式評估器，供月資金計劃表「線上表單」即時顯示公式欄。

只涵蓋資金計劃表會用到的語法：
  * 四則運算、比較、括號、字串連接 &
  * 同分頁 / 跨分頁儲存格參照（L4、$L$4、'2026(8)'!L60）與範圍（F3:F40）
  * 表格結構化參照（[@支出]、[@[提前扣款]]、表1[@支出]、表1[支出]）
  * 函數：SUM、IF、IFERROR、ABS、ROUND、MAX、MIN、AND、OR、NOT、N、COUNT、AVERAGE

無法評估時回傳 FormulaError，由呼叫端改用檔案內的快取值或留白，不臆測。
"""
from __future__ import annotations

import datetime as dt
import re
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple


class FormulaError(Exception):
    """公式無法評估（語法不支援、參照不存在、循環參照）。"""


# ---------------------------------------------------------------------------
# 詞法分析
# ---------------------------------------------------------------------------
_TOKEN_RE = re.compile(
    r"""
    (?P<ws>\s+)
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<sref>(?:[A-Za-z_一-鿿][\w一-鿿.]*)?\[(?:[^\[\]]|\[[^\[\]]*\])*\])
  | (?P<xref>(?:'(?:[^']|'')+'|[A-Za-z_一-鿿][\w一-鿿.()]*)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)
  | (?P<ref>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)
  | (?P<func>[A-Za-z_][A-Za-z0-9_.]*\s*\()
  | (?P<bool>TRUE|FALSE)
  | (?P<op><>|<=|>=|[-+*/^&=<>(),;%])
    """,
    re.X,
)


def tokenize(src: str) -> List[Tuple[str, str]]:
    out: List[Tuple[str, str]] = []
    pos = 0
    while pos < len(src):
        m = _TOKEN_RE.match(src, pos)
        if not m:
            raise FormulaError(f"無法解析: {src[pos:pos+12]!r}")
        pos = m.end()
        kind = m.lastgroup
        if kind == "ws":
            continue
        text = m.group(kind)
        if kind == "func":
            text = text[:-1].strip().upper()
        out.append((kind, text))
    return out


# ---------------------------------------------------------------------------
# 解析為 AST（tuple 形式）
# ---------------------------------------------------------------------------
class _Parser:
    def __init__(self, tokens: List[Tuple[str, str]]):
        self.t = tokens
        self.i = 0

    def peek(self, kind: Optional[str] = None, text: Optional[str] = None):
        if self.i >= len(self.t):
            return None
        k, v = self.t[self.i]
        if kind and k != kind:
            return None
        if text is not None and v != text:
            return None
        return self.t[self.i]

    def take(self, kind: Optional[str] = None, text: Optional[str] = None):
        tok = self.peek(kind, text)
        if tok is None:
            exp = text or kind or "token"
            got = self.t[self.i] if self.i < len(self.t) else "結尾"
            raise FormulaError(f"預期 {exp}，得到 {got}")
        self.i += 1
        return tok

    def parse(self):
        node = self.expr_cmp()
        if self.i != len(self.t):
            raise FormulaError(f"多餘的 token: {self.t[self.i]}")
        return node

    def expr_cmp(self):
        left = self.expr_concat()
        while self.peek("op") and self.peek()[1] in ("=", "<>", "<", "<=", ">", ">="):
            op = self.take()[1]
            right = self.expr_concat()
            left = ("cmp", op, left, right)
        return left

    def expr_concat(self):
        left = self.expr_add()
        while self.peek("op", "&"):
            self.take()
            right = self.expr_add()
            left = ("concat", left, right)
        return left

    def expr_add(self):
        left = self.expr_mul()
        while self.peek("op") and self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            right = self.expr_mul()
            left = ("bin", op, left, right)
        return left

    def expr_mul(self):
        left = self.expr_pow()
        while self.peek("op") and self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            right = self.expr_pow()
            left = ("bin", op, left, right)
        return left

    def expr_pow(self):
        left = self.expr_unary()
        while self.peek("op", "^"):
            self.take()
            right = self.expr_unary()
            left = ("bin", "^", left, right)
        return left

    def expr_unary(self):
        if self.peek("op", "-"):
            self.take()
            return ("neg", self.expr_unary())
        if self.peek("op", "+"):
            self.take()
            return self.expr_unary()
        node = self.expr_atom()
        while self.peek("op", "%"):
            self.take()
            node = ("bin", "/", node, ("num", 100.0))
        return node

    def expr_atom(self):
        tok = self.peek()
        if tok is None:
            raise FormulaError("公式不完整")
        kind, text = tok
        if kind == "num":
            self.take()
            return ("num", float(text))
        if kind == "str":
            self.take()
            return ("str", text[1:-1].replace('""', '"'))
        if kind == "bool":
            self.take()
            return ("bool", text == "TRUE")
        if kind == "ref":
            self.take()
            return ("ref", None, text.replace("$", ""))
        if kind == "xref":
            self.take()
            sheet, ref = text.rsplit("!", 1)
            if sheet.startswith("'"):
                sheet = sheet[1:-1].replace("''", "'")
            return ("ref", sheet, ref.replace("$", ""))
        if kind == "sref":
            self.take()
            return ("sref", text)
        if kind == "func":
            self.take()
            args = []
            if not self.peek("op", ")"):
                args.append(self.expr_cmp())
                while self.peek("op") and self.peek()[1] in (",", ";"):
                    self.take()
                    args.append(self.expr_cmp())
            self.take("op", ")")
            return ("call", text, args)
        if kind == "op" and text == "(":
            self.take()
            node = self.expr_cmp()
            self.take("op", ")")
            return node
        raise FormulaError(f"無法解析 token {tok}")


def parse(formula: str):
    src = formula[1:] if formula.startswith("=") else formula
    return _Parser(tokenize(src)).parse()


# ---------------------------------------------------------------------------
# 評估
# ---------------------------------------------------------------------------
_CELL_RE = re.compile(r"^([A-Za-z]{1,3})(\d+)$")


def _col_to_idx(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _idx_to_col(idx: int) -> str:
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _num(v: Any) -> float:
    """Excel 式數值轉換：空白→0、布林→0/1、日期→序號。字串盡量轉數字，否則 #VALUE!。"""
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, dt.datetime):
        return float((v - dt.datetime(1899, 12, 30)).days) + v.hour / 24 + v.minute / 1440
    if isinstance(v, dt.date):
        return float((v - dt.date(1899, 12, 30)).days)
    if isinstance(v, str):
        try:
            return float(v.replace(",", ""))
        except ValueError:
            raise FormulaError("#VALUE!")
    raise FormulaError("#VALUE!")


def _truthy(v: Any) -> bool:
    if isinstance(v, str):
        if v == "":
            return False
        if v.upper() == "TRUE":
            return True
        if v.upper() == "FALSE":
            return False
        raise FormulaError("#VALUE!")
    return bool(_num(v))


def _flatten(values: Sequence[Any]) -> List[Any]:
    out: List[Any] = []
    for v in values:
        if isinstance(v, list):
            out.extend(_flatten(v))
        else:
            out.append(v)
    return out


def _numeric_only(values: Sequence[Any]) -> List[float]:
    """SUM/AVERAGE 等對範圍的行為：忽略文字與空白，只取數值。"""
    out = []
    for v in _flatten(values):
        if isinstance(v, bool) or v is None or v == "":
            continue
        if isinstance(v, (int, float)):
            out.append(float(v))
        elif isinstance(v, (dt.date, dt.datetime)):
            out.append(_num(v))
    return out


def _cmp(op: str, a: Any, b: Any) -> bool:
    if isinstance(a, str) or isinstance(b, str):
        sa = "" if a is None else str(a)
        sb = "" if b is None else str(b)
        if not isinstance(a, str) or not isinstance(b, str):
            # 文字 vs 數字：Excel 視文字永遠大於數字；相等一律 False
            if op == "=":
                return False
            if op == "<>":
                return True
            return isinstance(a, str) if op in (">", ">=") else isinstance(b, str)
        a, b = sa.lower(), sb.lower()
    else:
        a, b = _num(a), _num(b)
    return {"=": a == b, "<>": a != b, "<": a < b, "<=": a <= b,
            ">": a > b, ">=": a >= b}[op]


class Evaluator:
    """對單一分頁的公式評估。

    get_cell(sheet, ref) 必須回傳「儲存格原始值」（數字、字串、日期或以 = 開頭的公式字串）。
    resolve_structured(sref, current_row) 回傳 (sheet, ref 或 range) 供結構化參照使用。
    """

    def __init__(self,
                 get_cell: Callable[[Optional[str], str], Any],
                 resolve_structured: Optional[Callable[[str, int], Tuple[Optional[str], str]]] = None):
        self.get_cell = get_cell
        self.resolve_structured = resolve_structured
        self._cache: Dict[Tuple[Optional[str], str], Any] = {}
        self._stack: set = set()
        self._ast_cache: Dict[str, Any] = {}

    def invalidate(self) -> None:
        self._cache.clear()

    # -- 公開 --
    def value(self, sheet: Optional[str], ref: str) -> Any:
        """回傳儲存格「計算後」的值。"""
        key = (sheet, ref.upper())
        if key in self._cache:
            return self._cache[key]
        if key in self._stack:
            raise FormulaError("#CYCLE")
        raw = self.get_cell(sheet, ref.upper())
        if isinstance(raw, str) and raw.startswith("="):
            self._stack.add(key)
            try:
                m = _CELL_RE.match(ref.upper())
                row = int(m.group(2)) if m else 0
                val = self.eval_formula(raw, sheet, row)
            finally:
                self._stack.discard(key)
        else:
            val = raw
        self._cache[key] = val
        return val

    def eval_formula(self, formula: str, sheet: Optional[str], row: int) -> Any:
        ast = self._ast_cache.get(formula)
        if ast is None:
            ast = parse(formula)
            self._ast_cache[formula] = ast
        return self._ev(ast, sheet, row)

    # -- 內部 --
    def _range(self, sheet: Optional[str], ref: str) -> List[Any]:
        a, b = ref.split(":")
        ma, mb = _CELL_RE.match(a), _CELL_RE.match(b)
        if not (ma and mb):
            raise FormulaError(f"不支援的範圍 {ref}")
        c0, c1 = sorted((_col_to_idx(ma.group(1)), _col_to_idx(mb.group(1))))
        r0, r1 = sorted((int(ma.group(2)), int(mb.group(2))))
        if (r1 - r0 + 1) * (c1 - c0 + 1) > 20000:
            raise FormulaError("範圍過大")
        out = []
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                out.append(self.value(sheet, f"{_idx_to_col(c)}{r}"))
        return out

    def _ev(self, node, sheet, row):
        kind = node[0]
        if kind == "num":
            return node[1]
        if kind == "str":
            return node[1]
        if kind == "bool":
            return node[1]
        if kind == "ref":
            s = node[1] if node[1] is not None else sheet
            if ":" in node[2]:
                return self._range(s, node[2])
            return self.value(s, node[2])
        if kind == "sref":
            if not self.resolve_structured:
                raise FormulaError("不支援結構化參照")
            s, ref = self.resolve_structured(node[1], row)
            s = s if s is not None else sheet
            if ":" in ref:
                return self._range(s, ref)
            return self.value(s, ref)
        if kind == "neg":
            return -_num(self._ev(node[1], sheet, row))
        if kind == "bin":
            op = node[1]
            a = _num(self._ev(node[2], sheet, row))
            b = _num(self._ev(node[3], sheet, row))
            if op == "+":
                return a + b
            if op == "-":
                return a - b
            if op == "*":
                return a * b
            if op == "/":
                if b == 0:
                    raise FormulaError("#DIV/0!")
                return a / b
            if op == "^":
                return a ** b
        if kind == "concat":
            a = self._ev(node[1], sheet, row)
            b = self._ev(node[2], sheet, row)
            return _to_text(a) + _to_text(b)
        if kind == "cmp":
            return _cmp(node[1], self._ev(node[2], sheet, row), self._ev(node[3], sheet, row))
        if kind == "call":
            return self._call(node[1], node[2], sheet, row)
        raise FormulaError(f"未知節點 {kind}")

    def _call(self, name: str, args, sheet, row):
        ev = lambda n: self._ev(n, sheet, row)  # noqa: E731
        if name == "IF":
            cond = _truthy(ev(args[0]))
            if cond:
                return ev(args[1]) if len(args) > 1 else True
            return ev(args[2]) if len(args) > 2 else False
        if name == "IFERROR":
            try:
                return ev(args[0])
            except FormulaError:
                return ev(args[1]) if len(args) > 1 else ""
        if name == "AND":
            return all(_truthy(v) for v in _flatten([ev(a) for a in args]))
        if name == "OR":
            return any(_truthy(v) for v in _flatten([ev(a) for a in args]))
        if name == "NOT":
            return not _truthy(ev(args[0]))
        vals = [ev(a) for a in args]
        if name == "SUM":
            return sum(_numeric_only(vals))
        if name == "COUNT":
            return float(len(_numeric_only(vals)))
        if name == "AVERAGE":
            nums = _numeric_only(vals)
            if not nums:
                raise FormulaError("#DIV/0!")
            return sum(nums) / len(nums)
        if name == "MAX":
            nums = _numeric_only(vals)
            return max(nums) if nums else 0.0
        if name == "MIN":
            nums = _numeric_only(vals)
            return min(nums) if nums else 0.0
        if name == "ABS":
            return abs(_num(vals[0]))
        if name == "ROUND":
            digits = int(_num(vals[1])) if len(vals) > 1 else 0
            x = _num(vals[0])
            # Excel 四捨五入（遠離零），非 Python 的銀行家捨入
            q = 10 ** digits
            return float(int(abs(x) * q + 0.5 + 1e-9) / q) * (1 if x >= 0 else -1)
        if name == "N":
            v = vals[0]
            return _num(v) if not isinstance(v, str) else 0.0
        if name == "TEXT":
            return excel_text(vals[0], str(vals[1]) if len(vals) > 1 else "General")
        if name == "TODAY":
            return dt.datetime.combine(dt.date.today(), dt.time())
        if name == "YEAR":
            return float(_to_date(vals[0]).year)
        if name == "MONTH":
            return float(_to_date(vals[0]).month)
        if name == "DAY":
            return float(_to_date(vals[0]).day)
        if name == "WEEKDAY":
            d = _to_date(vals[0])
            typ = int(_num(vals[1])) if len(vals) > 1 else 1
            wd = d.weekday()  # Mon=0
            if typ == 1:
                return float((wd + 1) % 7 + 1)
            if typ == 2:
                return float(wd + 1)
            if typ == 3:
                return float(wd)
            return float((wd + 1) % 7 + 1)
        if name == "DATE":
            return dt.datetime(int(_num(vals[0])), int(_num(vals[1])), int(_num(vals[2])))
        if name == "TRUE":
            return True
        if name == "FALSE":
            return False
        raise FormulaError(f"不支援的函數 {name}")


_EPOCH = dt.datetime(1899, 12, 30)
_WEEK_ZH = ["一", "二", "三", "四", "五", "六", "日"]
_WEEK_EN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
_MON_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _to_date(v: Any) -> dt.datetime:
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime.combine(v, dt.time())
    n = _num(v)
    return _EPOCH + dt.timedelta(days=n)


def excel_text(v: Any, fmt: str) -> str:
    """TEXT() 的常用子集：民國/西元日期、星期（aaa/aaaa/ddd/dddd）、千分位數字。"""
    f = fmt.strip()
    if f.startswith("[$-") and "]" in f:
        f = f[f.index("]") + 1:]
    low = f.lower()
    is_date = any(t in low for t in ("y", "e", "d", "a", "m")) and not any(ch in low for ch in "#0?")
    if is_date:
        if v in (None, ""):
            return ""
        d = _to_date(v)
        out = ""
        i = 0
        while i < len(f):
            ch = f[i]
            j = i
            while j < len(f) and f[j].lower() == ch.lower():
                j += 1
            run, n = f[i:j], j - i
            c = ch.lower()
            if c == "y":
                out += str(d.year) if n >= 3 else f"{d.year % 100:02d}"
            elif c == "e":
                out += str(d.year - 1911)
            elif c == "m":
                out += (_MON_EN[d.month - 1] if n == 3 else (f"{d.month:02d}" if n == 2 else str(d.month)))
            elif c == "d":
                if n >= 4:
                    out += ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][d.weekday()]
                elif n == 3:
                    out += _WEEK_EN[d.weekday()]
                else:
                    out += f"{d.day:02d}" if n == 2 else str(d.day)
            elif c == "a":
                out += ("星期" if n >= 4 else "") + _WEEK_ZH[d.weekday()]
            else:
                out += run.replace('"', "")
            i = j
        return out
    n = _num(v)
    if "," in f:
        dec = len(f.split(".")[1].rstrip('"')) if "." in f else 0
        return f"{n:,.{dec}f}"
    if "." in f:
        dec = len(f.split(".")[1])
        return f"{n:.{dec}f}"
    if f in ("0", "#", "General", ""):
        return _to_text(n if not n.is_integer() else float(int(n)))
    return _to_text(n)


def _to_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)
