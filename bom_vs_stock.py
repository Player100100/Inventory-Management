#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOM 与库存表比对工具

用法:
    python bom_vs_stock.py BOM.xlsx 库存.xlsx
    python bom_vs_stock.py BOM.xlsx 库存.xlsx -o 结果.xlsx
    python bom_vs_stock.py BOM.xlsx 库存.xlsx --slim          # 精简：删掉库存已提供的 BOM 列
    python bom_vs_stock.py BOM.xlsx 库存.xlsx --threshold 0.8 # 调模糊匹配阈值

输出工作簿包含三张表:
    比对结果  BOM 每一行 + 匹配到的库存信息 + 缺口
    待采购    库存不足 / 库存里没有的物料
    未匹配    BOM 行在库存里完全找不到对应货品

匹配优先级:
    1. 立创编号完全一致（BOM Supplier Part / Device 里的 _Cxxxxxx  ↔  库存「立创商城商品编号」）
    2. 型号名称归一化后完全一致（Device / Manufacturer Part / Comment ↔「货品」）
    3. 名称包含关系（短的那个至少 4 个字符）
    4. 相似度匹配（difflib，默认阈值 0.82，可用 --threshold 调整）

贴片阻容免购:
    库存表里查不到时，若该行是 0603 封装的电阻/电容，且阻值/容值出现在
    电阻电容.md 里（--rc-list 可指定，--no-rc 关闭），则判为「自有(免购)」，
    不计入待采购。值会先折算到欧姆 / pF 再比对。
"""
from __future__ import annotations

import argparse
import difflib
import re
import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 常量

BOM_HEADER_HINTS = ("quantity", "comment", "designator", "supplier part",
                    "manufacturer part", "device", "footprint")
STOCK_HEADER_HINTS = ("货品", "货架", "现有库存数量", "货品类型")

# BOM 列 → 库存列：语义上重复的列（--slim 时按这个映射删除）
REDUNDANT_COLUMNS = {
    "Manufacturer Part": "货品",
    "Device": "货品",
    "Comment": "货品",
    "Supplier Part": "立创商城商品编号",
    "Value": "参数简介",
}

STATUS_OK = "库存充足"
STATUS_SHORT = "库存不足"
STATUS_MISSING = "库存无此货"
STATUS_UNKNOWN = "库存未知"
STATUS_OWN = "自有(免购)"

# ---------------------------------------------------------------- 文本归一化


def norm(v: Any) -> str:
    """去空白、全角转半角、统一大写，用于比较。"""
    if v is None:
        return ""
    s = str(v)
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        elif o == 0x3000:
            out.append(" ")
        elif ch in "\r\n":
            out.append(" ")
        else:
            out.append(ch)
    s = "".join(out).upper()
    return re.sub(r"\s+", "", s)


def norm_name(v: Any) -> str:
    """比 norm 更宽松：仅用于名称相似度，去掉常见分隔符。"""
    s = norm(v)
    return re.sub(r"[/_\-\s.,()（）]", "", s)


def extract_lcsc(v: Any) -> str | None:
    """从单元格里抠出立创编号，如 'AO3400A_C20917' -> 'C20917'。"""
    s = str(v or "").strip()
    m = re.search(r"(?<![0-9A-Za-z])(C\d{4,9})(?![0-9A-Za-z])", s)
    return m.group(1) if m else None


def strip_device_suffix(v: Any) -> str:
    """'AO3400A_C20917' -> 'AO3400A'。"""
    s = str(v or "").strip()
    s = re.sub(r"[_\s]*C\d{4,9}$", "", s)
    return s.strip()


def to_num(v: Any) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        return None


# ---------------------------------------------------------------- 贴片阻容自有清单
# 电阻电容.md 记录了你料盒里有的贴片电阻/电容规格；0603 的阻容只要值在里面就不用买。

RC_MARK = re.compile(r"[（(]\s*单位\s*[:：]?\s*([^)）]+?)\s*[)）]\s*[:：]?")


def unit_factor(unit: str, kind: str) -> float | None:
    u = unit.strip()
    if kind == "R":          # 基准单位：欧姆
        if u in ("欧姆", "Ω", "ω", "欧", "R", "r"):
            return 1.0
        if u in ("K欧姆", "kΩ", "KΩ", "k", "K", "千欧", "k欧"):
            return 1e3
        if u in ("M欧姆", "MΩ", "M", "兆欧", "M欧"):
            return 1e6
        if u in ("mΩ", "m欧", "毫欧", "m"):
            return 1e-3
    else:                    # 基准单位：pF
        if u in ("pF", "pf", "PF", "p"):
            return 1.0
        if u in ("nF", "nf", "NF", "n"):
            return 1e3
        if u in ("uF", "µF", "μF", "u", "µ", "μ", "UF"):
            return 1e6
        if u == "F":
            return 1e12
    return None


class RcInventory:
    """电阻电容.md 解析结果，值统一折算到欧姆 / pF。"""

    def __init__(self):
        self.values: dict[str, list[float]] = {"R": [], "C": []}

    @classmethod
    def load(cls, path: Path) -> "RcInventory":
        inv = cls()
        text = path.read_text(encoding="utf-8")
        for line in text.splitlines():
            s = line.strip()
            if not s:
                continue
            if s.startswith("电阻"):
                kind = "R"
            elif s.startswith("电容"):
                kind = "C"
            else:
                continue
            unit, pos = None, 0
            for m in RC_MARK.finditer(s):
                inv._add(kind, unit, s[pos:m.start()])
                unit, pos = m.group(1).strip(), m.end()
            inv._add(kind, unit, s[pos:])
        return inv

    def _add(self, kind: str, unit: str | None, seg: str) -> None:
        if unit is None:
            return
        factor = unit_factor(unit, kind)
        if factor is None:
            return
        for tok in re.split(r"[,，;；]", seg):
            tok = tok.strip()
            if not tok:
                continue
            try:
                v = float(tok)
            except ValueError:
                continue
            self.values[kind].append(round(v * factor, 6))

    def has(self, kind: str, base_value: float) -> bool:
        for x in self.values[kind]:
            if abs(x - base_value) <= 1e-6 * max(1.0, abs(x)):
                return True
        return False

    def __bool__(self):
        return bool(self.values["R"] or self.values["C"])


BOM_VALUE_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*([A-Za-zμµΩ]*)\s*$")


def parse_value(v: Any, kind: str) -> float | None:
    """'10kΩ' -> 10000.0（欧姆）；'100nF' -> 100000.0（pF）。"""
    m = BOM_VALUE_RE.match(str(v or ""))
    if not m:
        return None
    num = float(m.group(1))
    suf = m.group(2)
    factor = unit_factor(suf, kind) if suf else 1.0
    if factor is None:
        return None
    return num * factor


def rc_own_check(rec: dict, cols: dict, inv: RcInventory) -> str | None:
    """0603 贴片阻容且值在清单里 -> 返回描述文字，否则 None。"""
    fp = str(rec.get(cols.get("Footprint") or "") or "").strip()
    if "0603" not in fp.upper():
        return None
    cat = str(rec.get(cols.get("Category") or "") or "")
    if "电容" in cat:
        kind = "C"
    elif "电阻" in cat:
        kind = "R"
    else:
        head = fp[:1].upper()
        kind = "C" if head == "C" else "R" if head == "R" else ""
    if not kind:
        return None
    text = ""
    for key in ("Value", "Comment"):
        h = cols.get(key)
        if h and rec.get(h) not in (None, ""):
            text = str(rec.get(h))
            break
    base = parse_value(text, kind)
    if base is None:
        return None
    return f"{text} {fp}" if inv.has(kind, base) else None


# ---------------------------------------------------------------- 读表


def read_sheet(path: Path, hints: tuple[str, ...]) -> tuple[list[str], list[dict]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hint_set = [norm(h) for h in hints]

    header_idx = None
    for i, row in enumerate(rows[:15]):
        vals = [norm(c) for c in row if c is not None]
        hit = sum(1 for h in hint_set if any(h in v for v in vals))
        if hit >= 2:
            header_idx = i
            break
    if header_idx is None:
        raise SystemExit(f"[错误] {path} 前 15 行里找不到表头")

    raw_header = rows[header_idx]
    header: list[str] = []
    for i, c in enumerate(raw_header):
        name = str(c).strip() if c is not None else ""
        header.append(name or f"列{i + 1}")

    data = []
    for row in rows[header_idx + 1:]:
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        rec = {header[i]: row[i] if i < len(row) else None for i in range(len(header))}
        data.append(rec)
    return header, data


def resolve(header: list[str], aliases: tuple[str, ...]) -> str | None:
    for a in aliases:
        na = norm(a)
        for h in header:
            if norm(h) == na:
                return h
    for a in aliases:
        na = norm(a)
        for h in header:
            if na and na in norm(h):
                return h
    return None


# ---------------------------------------------------------------- 匹配


class StockItem:
    __slots__ = ("row", "name", "raw_name", "kind", "shelf", "qty", "lcsc", "param", "price")

    def __init__(self, row: dict, cols: dict):
        self.row = row
        self.name = norm_name(row.get(cols["货品"]))
        self.raw_name = row.get(cols["货品"])
        self.kind = row.get(cols["货品类型"])
        self.shelf = row.get(cols["货架"])
        self.qty = to_num(row.get(cols["库存"]))
        self.lcsc = norm(row.get(cols["lcsc"]))
        self.param = row.get(cols["参数"])
        self.price = row.get(cols["单价"])


def match_score(bom_names: list[str], bom_codes: set[str],
                item: StockItem, threshold: float) -> tuple[float, str]:
    """返回 (分数 0~100, 匹配方式)；0 表示不匹配。"""
    if bom_codes and item.lcsc and item.lcsc in bom_codes:
        return 100.0, "立创编号"

    best = 0.0
    best_method = ""
    for raw in bom_names:
        a = norm_name(raw)
        if not a or not item.name:
            continue
        if a == item.name:
            return 95.0, "名称完全一致"
        short, long = (a, item.name) if len(a) <= len(item.name) else (item.name, a)
        if len(short) >= 4 and short in long:
            score = 85.0
            if score > best:
                best, best_method = score, "名称包含"
        r = difflib.SequenceMatcher(None, a, item.name).ratio()
        if r >= threshold:
            score = 60.0 + 25.0 * (r - threshold) / (1.0 - threshold)
            if score > best:
                best, best_method = score, f"相似度 {r:.0%}"
            continue
        # 前缀一致兜底：型号前 8 位相同（如 SPX1117M3-3.3 / SPX1117M3-L-3-3/TR(MS)），
        # 但相似度又不能太低，避免把 SMBJ30CA 配到 SMBJ36CA 这种电压档不同的料上。
        prefix = 0
        for x, y in zip(a, item.name):
            if x != y:
                break
            prefix += 1
        if prefix >= 8 and r >= 0.72:
            score = 60.0 + 15.0 * (r - 0.72) / 0.28
            if score > best:
                best, best_method = score, f"前缀一致 + 相似度 {r:.0%}"
    return best, best_method


def build_bom_view(rec: dict, cols: dict) -> tuple[list[str], set[str]]:
    names = []
    for key in ("Device", "ManufacturerPart", "Comment", "Value"):
        h = cols.get(key)
        if h:
            v = rec.get(h)
            if v not in (None, ""):
                names.append(strip_device_suffix(v) if key == "Device" else str(v))
    seen, uniq = set(), []
    for n in names:
        k = norm_name(n)
        if k and k not in seen:
            seen.add(k)
            uniq.append(n)
    codes = set()
    for key in ("SupplierPart", "Device"):
        h = cols.get(key)
        if h:
            c = extract_lcsc(rec.get(h))
            if c:
                codes.add(c)
    return uniq, codes


# ---------------------------------------------------------------- 输出


HEADER_FILL = "DCE6F1"
SHORT_FILL = "FCE4D6"
MISS_FILL = "F8CBAD"
OK_FILL = "E2EFDA"
OWN_FILL = "DDEBF7"


def write_sheet(ws, header: list[str], rows: list[list], fills: list[str | None]):
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=HEADER_FILL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row, fill in zip(rows, fills):
        ws.append(row)
        if fill:
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", fgColor=fill)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    for i, h in enumerate(header, 1):
        width = max(8, min(42, len(str(h)) * 2 + 4))
        for r in range(1, min(ws.max_row, 200) + 1):
            v = ws.cell(r, i).value
            if v is not None:
                width = max(width, min(42, len(str(v)) * 1.6 + 2))
        ws.column_dimensions[get_column_letter(i)].width = width


# ---------------------------------------------------------------- 主流程


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="BOM 与库存表比对")
    ap.add_argument("bom", type=Path, help="BOM xlsx")
    ap.add_argument("stock", type=Path, help="库存 xlsx")
    ap.add_argument("-o", "--out", type=Path, default=None, help="输出 xlsx（默认 <BOM名>_比对结果.xlsx）")
    ap.add_argument("--threshold", type=float, default=0.82, help="模糊匹配阈值，默认 0.82")
    ap.add_argument("--slim", action="store_true",
                    help="删除库存表已提供的 BOM 列（仅当匹配行上库存侧确实有值时才删）")
    ap.add_argument("--rc-list", type=Path, default=None,
                    help="0603 贴片阻容自有清单 md（默认用脚本同目录的 电阻电容.md）")
    ap.add_argument("--no-rc", action="store_true", help="不使用 0603 贴片阻容自有清单")
    args = ap.parse_args(argv)
    if args.rc_list is None:
        args.rc_list = Path(__file__).with_name("电阻电容.md")
    if args.no_rc:
        args.rc_list = None

    for p in (args.bom, args.stock):
        if not p.exists():
            raise SystemExit(f"[错误] 找不到文件: {p}")

    bom_header, bom_rows = read_sheet(args.bom, BOM_HEADER_HINTS)
    stock_header, stock_rows = read_sheet(args.stock, STOCK_HEADER_HINTS)

    # ---- 定位列
    bom_cols = {
        "Quantity": resolve(bom_header, ("Quantity", "数量")),
        "Designator": resolve(bom_header, ("Designator", "位号")),
        "Comment": resolve(bom_header, ("Comment", "名称")),
        "Value": resolve(bom_header, ("Value", "值")),
        "Footprint": resolve(bom_header, ("Footprint", "封装")),
        "ManufacturerPart": resolve(bom_header, ("Manufacturer Part", "厂商型号")),
        "Manufacturer": resolve(bom_header, ("Manufacturer", "厂商")),
        "SupplierPart": resolve(bom_header, ("Supplier Part", "供应商编号")),
        "Device": resolve(bom_header, ("Device", "器件", "器件名称")),
        "Category": resolve(bom_header, ("Category", "分类")),
    }
    stock_cols = {
        "货品": resolve(stock_header, ("货品", "名称", "型号")),
        "货架": resolve(stock_header, ("货架", "位置")),
        "货品类型": resolve(stock_header, ("货品类型", "类型")),
        "单价": resolve(stock_header, ("单价", "价格")),
        "库存": resolve(stock_header, ("现有库存数量", "现有库存", "库存")),
        "lcsc": resolve(stock_header, ("立创商城商品编号", "商品编号", "立创编号")),
        "参数": resolve(stock_header, ("参数简介", "参数")),
    }
    if not stock_cols["货品"] or not stock_cols["库存"]:
        raise SystemExit("[错误] 库存表缺少「货品」或「现有库存数量」列")
    if not bom_cols["Quantity"]:
        raise SystemExit("[错误] BOM 缺少 Quantity 列")

    items = [StockItem(r, stock_cols) for r in stock_rows
             if (r.get(stock_cols["货品"]) not in (None, ""))]
    items = [it for it in items if it.name]

    rc_inv = None
    if args.rc_list:
        if args.rc_list.exists():
            rc_inv = RcInventory.load(args.rc_list)
        else:
            print(f"[提示] 没找到阻容清单 {args.rc_list}，跳过 0603 阻容免购判断")

    # ---- 逐行比对
    detail_rows: list[list] = []
    detail_fills: list[str | None] = []
    to_buy: list[list] = []
    unmatched: list[list] = []
    stats = {STATUS_OK: 0, STATUS_OWN: 0, STATUS_SHORT: 0, STATUS_MISSING: 0, STATUS_UNKNOWN: 0}

    # 精简模式的判定素材：匹配行上库存侧是否真有值
    redundant_probe = {c: [0, 0] for c in REDUNDANT_COLUMNS
                       if c in bom_header and REDUNDANT_COLUMNS[c] in stock_header}

    matched_pairs = []  # (bom_row_values, stock_item)

    for rec in bom_rows:
        names, codes = build_bom_view(rec, bom_cols)
        best_item, best_score, best_method = None, 0.0, ""
        for it in items:
            score, method = match_score(names, codes, it, args.threshold)
            if score > best_score:
                best_item, best_score, best_method = it, score, method

        # 库存表里没有时，再看 0603 贴片阻容自有清单
        own_label = None
        if best_item is None and rc_inv is not None:
            own_label = rc_own_check(rec, bom_cols, rc_inv)

        need = to_num(rec.get(bom_cols["Quantity"])) or 0
        if best_item is None and own_label:
            status = STATUS_OWN
            stock_qty_disp = "免购"
            gap = 0
        elif best_item is None:
            status = STATUS_MISSING
            stock_qty_disp = ""
            gap = need
        else:
            matched_pairs.append((rec, best_item))
            for col in redundant_probe:
                val = rec.get(col)
                if val not in (None, ""):
                    redundant_probe[col][1] += 1
                    src = best_item.row.get(stock_cols.get(REDUNDANT_COLUMNS[col], ""), None)
                    if src not in (None, ""):
                        redundant_probe[col][0] += 1
            if best_item.qty is None:
                status, stock_qty_disp, gap = STATUS_UNKNOWN, "未知", None
            elif best_item.qty < 0:
                status, stock_qty_disp, gap = STATUS_OK, "充足(-1)", 0
            elif best_item.qty >= need:
                status, stock_qty_disp, gap = STATUS_OK, int(best_item.qty), 0
            else:
                status = STATUS_SHORT
                stock_qty_disp = int(best_item.qty)
                gap = need - best_item.qty
        stats[status] += 1

        detail_rows.append([
            rec.get(bom_cols["Designator"]),
            rec.get(bom_cols["Comment"]),
            rec.get(bom_cols["Value"]),
            rec.get(bom_cols["Footprint"]),
            rec.get(bom_cols["ManufacturerPart"]),
            rec.get(bom_cols["SupplierPart"]),
            int(need) if need == int(need) else need,
            "0603阻容清单" if own_label else (best_method if best_item else ""),
            own_label or (best_item.raw_name if best_item else ""),
            best_item.shelf if best_item else "",
            "0603阻容" if own_label else (best_item.kind if best_item else ""),
            stock_qty_disp,
            "" if gap is None else (int(gap) if gap == int(gap) else gap),
            status,
        ])
        detail_fills.append(
            OK_FILL if status == STATUS_OK else
            OWN_FILL if status == STATUS_OWN else
            SHORT_FILL if status == STATUS_SHORT else
            MISS_FILL if status == STATUS_MISSING else None
        )

        purchase = [rec.get(bom_cols["Designator"]), rec.get(bom_cols["Comment"]),
                    rec.get(bom_cols["ManufacturerPart"]), rec.get(bom_cols["SupplierPart"]),
                    int(need) if need == int(need) else need,
                    stock_qty_disp,
                    "" if gap is None else (int(gap) if gap == int(gap) else gap),
                    status]
        if status in (STATUS_SHORT, STATUS_MISSING, STATUS_UNKNOWN):
            to_buy.append(purchase)
        if status == STATUS_MISSING:
            unmatched.append(purchase)

    detail_header = ["位号", "名称", "值", "封装", "厂商型号", "立创编号",
                     "需求数量", "匹配方式", "库存货品", "货架", "货品类型",
                     "现有库存", "缺口", "结论"]
    buy_header = ["位号", "名称", "厂商型号", "立创编号", "需求数量",
                  "现有库存", "缺口", "结论"]

    # ---- 精简模式
    # 规则：只有当「BOM 里这一列的所有非空行」都已在库存里匹配上、且库存侧对应值也齐全时，
    # 才把这一列从 BOM 里删掉。未匹配的行是你要去买的，删了型号就没法买了。
    keep_header = list(detail_header)
    slim_header: list[str] = []
    slim_rows: list[list] = []
    drop_report = []
    if args.slim:
        value_rows = {}   # 列名 -> [(BOM值, 是否匹配上, 库存侧是否有值)]
        for rec, it in matched_pairs:
            for col in redundant_probe:
                v = rec.get(col)
                if v not in (None, ""):
                    src = it.row.get(stock_cols.get(REDUNDANT_COLUMNS[col], ""), None)
                    value_rows.setdefault(col, []).append((v, True, src not in (None, "")))
        matched_set = {id(rec) for rec, _ in matched_pairs}
        for rec in bom_rows:
            if id(rec) in matched_set:
                continue
            for col in redundant_probe:
                v = rec.get(col)
                if v not in (None, ""):
                    value_rows.setdefault(col, []).append((v, False, False))

        dropped = set()
        for col in REDUNDANT_COLUMNS:
            if col not in value_rows:
                continue
            rows_ = value_rows[col]
            unmatched_n = sum(1 for _, m, _ in rows_ if not m)
            covered = sum(1 for _, m, has in rows_ if m and has)
            matched_n = sum(1 for _, m, _ in rows_ if m)
            if unmatched_n == 0 and matched_n and covered == matched_n:
                dropped.add(col)
                drop_report.append((col, covered, len(rows_), "已删除"))
            else:
                why = f"有 {unmatched_n} 行库存里没有" if unmatched_n else \
                      f"库存侧仅 {covered}/{matched_n} 有值"
                drop_report.append((col, covered, len(rows_), f"保留（{why}）"))

        slim_header = [h for h in bom_header if h not in dropped]
        slim_rows = [[rec.get(h) for h in slim_header] for rec in bom_rows]

    # ---- 写文件
    out = args.out or args.bom.with_name(args.bom.stem + "_比对结果.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "比对结果"
    write_sheet(ws, keep_header, detail_rows, detail_fills)
    ws2 = wb.create_sheet("待采购")
    write_sheet(ws2, buy_header, to_buy, [SHORT_FILL] * len(to_buy))
    ws3 = wb.create_sheet("未匹配")
    write_sheet(ws3, buy_header, unmatched, [MISS_FILL] * len(unmatched))
    if args.slim:
        ws4 = wb.create_sheet("精简BOM")
        write_sheet(ws4, slim_header, slim_rows, [None] * len(slim_rows))
    wb.save(out)

    # ---- 控制台报告
    print(f"BOM    : {args.bom.name}  ({len(bom_rows)} 行)")
    print(f"库存表 : {args.stock.name}  ({len(items)} 条货品)")
    print(f"输出   : {out}")
    print()
    print(f"  库存充足   : {stats[STATUS_OK]}")
    print(f"  自有(免购) : {stats[STATUS_OWN]}")
    print(f"  库存不足   : {stats[STATUS_SHORT]}")
    print(f"  库存无此货 : {stats[STATUS_MISSING]}")
    print(f"  库存未知   : {stats[STATUS_UNKNOWN]}")
    print()
    print("匹配到的物料：")
    for rec, it in matched_pairs:
        print(f"  {str(rec.get(bom_cols['Comment'])):<24} -> {it.raw_name}  "
              f"货架 {it.shelf}  库存 {it.qty}")
    if drop_report:
        print()
        print("精简模式列处理：")
        for col, have, total, verdict in drop_report:
            print(f"  {col:<18} 库存侧有值 {have}/{total}  {verdict}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
