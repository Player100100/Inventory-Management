#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BOM 采购计算器

图形界面：选择 BOM 文件、输入生产套数、选择库存表
  -> 需求数量 = BOM Quantity × 套数
  -> 与库存逐项比对，扣减用掉的量（自动备份库存表）
  -> 输出「删掉已满足项」的需采购 BOM（Quantity 列 = 实际要买的缺口）
  -> 可选：运行前 git pull、扣减后 commit + push 库存表，多人共用一份库存

BOM 必填 Quantity；Device / Manufacturer Part / Comment 至少一个；
Supplier Part、Footprint、Value、Category、Designator 建议保留。
详见 README.md 或界面上的「BOM 要求」按钮。

命令行（便于脚本化 / 测试）:
    python bom_gui.pyw --cli BOM.xlsx --sets 10
    python bom_gui.pyw --cli BOM.xlsx --sets 10 --stock 库存.xlsx -o 采购.xlsx
    python bom_gui.pyw --cli BOM.xlsx --sets 10 --dry-run     # 只算不扣库存
    python bom_gui.pyw --cli BOM.xlsx --sets 10 --no-sync     # 不同步 GitHub
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

from bom_vs_stock import (
    BOM_HEADER_HINTS,
    STOCK_HEADER_HINTS,
    STATUS_MISSING,
    STATUS_OK,
    STATUS_OWN,
    STATUS_SHORT,
    STATUS_UNKNOWN,
    MISS_FILL,
    OK_FILL,
    OWN_FILL,
    SHORT_FILL,
    RcInventory,
    StockItem,
    build_bom_view,
    match_score,
    norm,
    read_sheet,
    rc_own_check,
    resolve,
    to_num,
    write_sheet,
)

def _base_dir() -> Path:
    """源码运行时是脚本目录；PyInstaller 打包后是 exe 所在目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = _base_dir()
DEFAULT_STOCK = BASE_DIR / "库存管理 副本.xlsx"
DEFAULT_RC = BASE_DIR / "电阻电容.md"


# ---------------------------------------------------------------- 工具


def _num(x):
    if x is None:
        return None
    f = round(float(x), 6)
    return int(f) if f.is_integer() else f


def _find_header_row(rows, hints, limit: int = 15) -> int:
    hint_set = [norm(h) for h in hints]
    for i, row in enumerate(rows[:limit]):
        vals = [norm(c) for c in row if c is not None]
        hit = sum(1 for h in hint_set if any(h in v for v in vals))
        if hit >= 2:
            return i
    raise ValueError("前 15 行里找不到表头")


# ---------------------------------------------------------------- GitHub 同步


class GitSyncError(RuntimeError):
    pass


class GitSync:
    """只同步库存 xlsx：运行前 pull，扣减后 add / commit / push。

    只操作指定的库存文件，不会把 BOM、采购结果等本地文件提交上去。
    """

    NET_ERR = ("Could not connect", "Failed to connect", "Connection refused",
               "Connection timed out", "unable to access",
               "Could not resolve host")

    def __init__(self, repo: Path):
        self.repo = Path(repo).resolve()
        self.git = shutil.which("git")
        self.available = bool(self.git) and (self.repo / ".git").exists()

    def _cmd(self, args, bypass_proxy: bool):
        cmd = ["git", "-C", str(self.repo)]
        if bypass_proxy:
            cmd[1:1] = ["-c", "http.proxy=", "-c", "https.proxy="]
        return cmd + args

    def _run(self, args, bypass_proxy: bool = False, timeout: int = 180):
        kwargs = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        return subprocess.run(
            self._cmd(args, bypass_proxy), capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout, **kwargs)

    @staticmethod
    def _msg(r) -> str:
        return (r.stderr or r.stdout or "").strip() or \
            f"git 退出码 {r.returncode}"

    def _is_net_err(self, r) -> bool:
        text = f"{r.stdout}\n{r.stderr}"
        return any(k in text for k in self.NET_ERR)

    def _run_net(self, args, timeout: int = 180):
        """先按全局配置走；代理连不上时自动绕过代理重试一次。"""
        r = self._run(args, timeout=timeout)
        if r.returncode != 0 and self._is_net_err(r):
            r2 = self._run(args, bypass_proxy=True, timeout=timeout)
            if r2.returncode == 0:
                return r2
            raise GitSyncError(self._msg(r2))
        return r

    def pull(self) -> str:
        r = self._run_net(["pull", "--ff-only"])
        if r.returncode != 0:
            raise GitSyncError(self._msg(r))
        return (r.stdout or "").strip()

    def commit_and_push(self, rel_path: str, message: str) -> str:
        """提交并推送单个文件；返回 pushed / nochange。"""
        r = self._run(["add", "--", rel_path])
        if r.returncode != 0:
            raise GitSyncError(self._msg(r))

        staged = self._run(["diff", "--cached", "--name-only"])
        if not (staged.stdout or "").strip():
            return "nochange"

        r = self._run(["commit", "-m", message])
        if r.returncode != 0:
            raise GitSyncError(self._msg(r))

        r = self._run_net(["push"])
        if r.returncode != 0:
            # 别人先推了：先 rebase 再重试
            pr = self._run_net(["pull", "--rebase"])
            if pr.returncode != 0:
                self._run(["rebase", "--abort"])
                raise GitSyncError(
                    "远程库存已被他人更新，自动合并失败。请先手动执行 "
                    "git pull 并解决冲突后再试。\n" + self._msg(pr))
            r = self._run_net(["push"])
            if r.returncode != 0:
                raise GitSyncError(self._msg(r))
        return "pushed"


BOM_HELP = """BOM 表要求（读取第一张工作表，表头需在前 15 行内）

【必填】
  Quantity（数量）
      每个位号的需求数量；最终需求 = 该值 × 生产套数。

【至少填一项，用于匹配库存】
  Device（器件）
  Manufacturer Part（厂商型号）
  Comment（名称）

【强烈建议】
  Supplier Part（供应商编号）
      如 C20917，立创编号匹配最准确。
  Footprint（封装）
      0603 电阻/电容免购判断需要。

【可选】
  Value（值）        如 10kΩ、100nF
  Category（分类）   区分电阻/电容
  Designator（位号） 输出与核对用
  Manufacturer（厂商）

【列名别名】
  数量、名称、位号、封装、值、分类、厂商型号、供应商编号 均可识别。
  立创 / 嘉立创 EDA 导出的 BOM 默认就包含以上列，可直接使用。

【库存表】
  必填：货品、现有库存数量
  可选：货架、货品类型、单价、立创商城商品编号、参数简介、总计消耗数量
  「现有库存数量」为 -1 时视为充足/不计数，不参与扣减。
"""


class StockBook:
    """库存工作簿：保留每个货品对应的 Excel 行号，便于扣减后写回。"""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.wb = openpyxl.load_workbook(self.path)
        self.ws = self.wb.worksheets[0]
        # 读值用 data_only 副本（公式取缓存结果）；写回仍用原工作簿，保留格式
        vals_wb = openpyxl.load_workbook(self.path, data_only=True)
        raw_rows = list(vals_wb.worksheets[0].iter_rows(values_only=True))
        idx = _find_header_row(raw_rows, STOCK_HEADER_HINTS)

        self.header = [
            str(c).strip() if c is not None else f"列{i + 1}"
            for i, c in enumerate(raw_rows[idx])
        ]
        self.cols = {
            "货品": resolve(self.header, ("货品", "名称", "型号")),
            "货架": resolve(self.header, ("货架", "位置")),
            "货品类型": resolve(self.header, ("货品类型", "类型")),
            "单价": resolve(self.header, ("单价", "价格")),
            "库存": resolve(self.header, ("现有库存数量", "现有库存", "库存")),
            "lcsc": resolve(self.header, ("立创商城商品编号", "商品编号", "立创编号")),
            "参数": resolve(self.header, ("参数简介", "参数")),
            "消耗": resolve(self.header, ("总计消耗数量", "消耗数量", "已消耗", "消耗")),
        }
        if not self.cols["货品"] or not self.cols["库存"]:
            raise ValueError("库存表缺少「货品」或「现有库存数量」列")

        # [(StockItem, Excel 行号, 原始记录)]
        self.entries: list[tuple[StockItem, int, dict]] = []
        for i in range(idx + 1, len(raw_rows)):
            values = raw_rows[i]
            rec = {
                self.header[c]: (values[c] if c < len(values) else None)
                for c in range(len(self.header))
            }
            if rec.get(self.cols["货品"]) in (None, ""):
                continue
            item = StockItem(rec, self.cols)
            if item.name:
                self.entries.append((item, i + 1, rec))


def _resolve_bom_cols(header: list[str]) -> dict:
    return {
        "Quantity": resolve(header, ("Quantity", "数量")),
        "Designator": resolve(header, ("Designator", "位号")),
        "Comment": resolve(header, ("Comment", "名称")),
        "Value": resolve(header, ("Value", "值")),
        "Footprint": resolve(header, ("Footprint", "封装")),
        "ManufacturerPart": resolve(header, ("Manufacturer Part", "厂商型号")),
        "Manufacturer": resolve(header, ("Manufacturer", "厂商")),
        "SupplierPart": resolve(header, ("Supplier Part", "供应商编号")),
        "Device": resolve(header, ("Device", "器件", "器件名称")),
        "Category": resolve(header, ("Category", "分类")),
    }


# ---------------------------------------------------------------- 核心计算


def compute(bom_path: Path, stock_path: Path, sets: int,
            rc_path: Path | None = DEFAULT_RC, threshold: float = 0.82) -> dict:
    """比对并计算缺口，不写任何文件。"""
    bom_header, bom_rows = read_sheet(Path(bom_path), BOM_HEADER_HINTS)
    cols = _resolve_bom_cols(bom_header)
    if not cols["Quantity"]:
        raise ValueError("BOM 缺少 Quantity 列")

    stock = StockBook(Path(stock_path))

    rc_inv = None
    if rc_path and Path(rc_path).exists():
        rc_inv = RcInventory.load(Path(rc_path))
        if not rc_inv:
            rc_inv = None

    # 库存余量（按对象跟踪，负库存/未知库存记为 None 单独处理）
    remaining: dict[int, float | None] = {}
    for item, _, _ in stock.entries:
        q = item.qty
        remaining[id(item)] = q if (q is not None and q >= 0) else None

    results = []
    used_by_item: dict[int, float] = {}
    stats = {STATUS_OK: 0, STATUS_OWN: 0, STATUS_SHORT: 0,
             STATUS_MISSING: 0, STATUS_UNKNOWN: 0}

    for rec in bom_rows:
        names, codes = build_bom_view(rec, cols)

        best_item, best_score, best_method = None, 0.0, ""
        for item, _, _ in stock.entries:
            score, method = match_score(names, codes, item, threshold)
            if score > best_score:
                best_item, best_score, best_method = item, score, method

        qty_raw = to_num(rec.get(cols["Quantity"])) or 0
        need = qty_raw * sets

        own_label = None
        if best_item is None and rc_inv is not None:
            own_label = rc_own_check(rec, cols, rc_inv)

        used = 0.0
        if best_item is None:
            if own_label:
                status, gap, avail = STATUS_OWN, 0.0, None
            else:
                status, gap, avail = STATUS_MISSING, need, None
        else:
            avail = remaining[id(best_item)]
            if avail is None:
                if best_item.qty is not None and best_item.qty < 0:
                    # 库存表里的 -1 视为「充足/不计数」
                    status, gap = STATUS_OK, 0.0
                else:
                    status, gap = STATUS_UNKNOWN, need
            else:
                used = min(need, avail)
                gap = need - used
                remaining[id(best_item)] = avail - used
                if used > 0:
                    used_by_item[id(best_item)] = \
                        used_by_item.get(id(best_item), 0.0) + used
                status = STATUS_OK if gap <= 0 else STATUS_SHORT

        stats[status] += 1
        results.append({
            "rec": rec,
            "need": need,
            "gap": gap,
            "used": used,
            "status": status,
            "item": best_item,
            "avail": avail,
            "own": own_label,
            "method": best_method,
        })

    # 需采购 BOM：只保留买得到的行，Quantity 换成实际缺口
    qty_col = cols["Quantity"]
    purchase_rows: list[tuple[list, str]] = []
    for r in results:
        if r["gap"] <= 0:
            continue
        row = [r["rec"].get(h) for h in bom_header]
        row[bom_header.index(qty_col)] = _num(r["gap"])
        purchase_rows.append((row, r["status"]))

    return {
        "bom_header": bom_header,
        "bom_rows": bom_rows,
        "cols": cols,
        "stock": stock,
        "results": results,
        "purchase_rows": purchase_rows,
        "used_by_item": used_by_item,
        "stats": stats,
        "sets": sets,
    }


def apply_stock(stock: StockBook, used_by_item: dict[int, float],
                backup: bool = True) -> tuple[list[tuple], Path | None]:
    """把用掉的量从库存表扣掉，同时累加「总计消耗数量」。"""
    cur_col = stock.cols["库存"]
    cur_idx = stock.header.index(cur_col) + 1
    cons_col = stock.cols.get("消耗")
    cons_idx = stock.header.index(cons_col) + 1 if cons_col else None

    changed = []
    for item, excel_row, rec in stock.entries:
        used = used_by_item.get(id(item))
        if not used or used <= 0:
            continue
        cur = to_num(rec.get(cur_col))
        if cur is None or cur < 0:
            continue
        new_cur = cur - used
        stock.ws.cell(excel_row, cur_idx).value = _num(new_cur)
        if cons_idx:
            cons = to_num(rec.get(cons_col)) or 0
            stock.ws.cell(excel_row, cons_idx).value = _num(cons + used)
        changed.append((item.raw_name, cur, new_cur, used))

    backup_path = None
    if changed and backup:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = stock.path.with_name(
            f"{stock.path.stem}_备份_{ts}{stock.path.suffix}")
        shutil.copy2(stock.path, backup_path)
    if changed:
        stock.wb.save(stock.path)
    return changed, backup_path


def run(bom_path, sets: int, stock_path=None, out_path=None,
        rc_path=DEFAULT_RC, threshold: float = 0.82,
        update_stock: bool = True, backup: bool = True,
        sync: bool = False, repo_dir=None, log=print) -> dict:
    bom_path = Path(bom_path)
    sets = int(sets)
    if sets <= 0:
        raise ValueError("套数必须是正整数")
    if not bom_path.exists():
        raise FileNotFoundError(f"找不到 BOM: {bom_path}")
    stock_path = Path(stock_path) if stock_path else DEFAULT_STOCK
    if not stock_path.exists():
        raise FileNotFoundError(f"找不到库存表: {stock_path}")

    repo_dir = Path(repo_dir).resolve() if repo_dir else BASE_DIR.resolve()

    # 先拉取最新库存，保证多人协作时基于最新数据扣减
    syncer = None
    if sync and update_stock:
        syncer = GitSync(repo_dir)
        if syncer.available:
            log("[同步] 正在拉取 GitHub 上的最新库存 ...")
            try:
                out = syncer.pull()
                first = next((ln for ln in out.splitlines()
                              if ln.strip() and "Already up to date" not in ln),
                             "已是最新")
                log(f"[同步] {first}")
            except (GitSyncError, subprocess.TimeoutExpired) as e:
                log(f"[同步] 拉取失败，继续使用本地库存：{e}")
        else:
            log("[同步] 未检测到 Git 仓库或 git 命令，跳过自动同步")
            syncer = None

    data = compute(bom_path, stock_path, sets, rc_path, threshold)
    stats = data["stats"]

    out_path = Path(out_path) if out_path else \
        bom_path.with_name(bom_path.stem + "_需采购.xlsx")
    if out_path.resolve() == bom_path.resolve():
        raise ValueError("输出文件不能覆盖 BOM 本身")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "需采购BOM"
    write_sheet(ws, data["bom_header"],
                [row for row, _ in data["purchase_rows"]],
                [None] * len(data["purchase_rows"]))

    # 附一张明细，方便核对每个缺口是怎么算出来的
    ws2 = wb.create_sheet("计算明细")
    cols = data["cols"]
    detail_rows, detail_fills = [], []
    for r in data["results"]:
        rec, item = r["rec"], r["item"]
        detail_rows.append([
            rec.get(cols["Designator"]),
            rec.get(cols["Comment"]),
            rec.get(cols["Footprint"]),
            _num(to_num(rec.get(cols["Quantity"])) or 0),
            sets,
            _num(r["need"]),
            r["own"] or (item.raw_name if item else ""),
            "0603阻容清单" if r["own"] else r["method"],
            _num(item.qty) if item else "",
            _num(r["avail"]) if r["avail"] is not None else "",
            _num(r["used"]),
            _num(r["gap"]),
            r["status"],
        ])
        detail_fills.append(
            OK_FILL if r["status"] == STATUS_OK else
            OWN_FILL if r["status"] == STATUS_OWN else
            SHORT_FILL if r["status"] == STATUS_SHORT else
            MISS_FILL if r["status"] == STATUS_MISSING else None)
    write_sheet(ws2, ["位号", "名称", "封装", "BOM数量", "套数", "需求数量",
                      "匹配货品", "匹配方式", "库存", "可用", "扣减", "缺口",
                      "结论"], detail_rows, detail_fills)

    wb.save(out_path)

    changed, backup_path = ([], None)
    if update_stock:
        changed, backup_path = apply_stock(
            data["stock"], data["used_by_item"], backup=backup)

    log(f"BOM      : {bom_path.name}  ({len(data['bom_rows'])} 行 × {sets} 套)")
    log(f"库存表   : {stock_path.name}  ({len(data['stock'].entries)} 条货品)")
    log(f"需采购   : {len(data['purchase_rows'])} 行 -> {out_path}")
    log("")
    log(f"  库存充足     : {stats[STATUS_OK]}")
    log(f"  自有(免购)   : {stats[STATUS_OWN]}")
    log(f"  库存不足     : {stats[STATUS_SHORT]}")
    log(f"  库存无此货   : {stats[STATUS_MISSING]}")
    log(f"  库存未知     : {stats[STATUS_UNKNOWN]}")
    if update_stock:
        log("")
        log(f"库存扣减   : {len(changed)} 种货品")
        for name, old, new, used in changed:
            log(f"  {str(name):<28} {_num(old)} -> {_num(new)}  (用 {_num(used)})")
        if backup_path:
            log(f"库存备份   : {backup_path.name}")
        if changed:
            log("提示：库存已实际扣减，同一批生产请勿重复运行；"
                "如需重算，请用备份恢复库存表。")
    else:
        log("")
        log("库存扣减   : 已跳过（dry-run）")

    # 扣减完成后把库存表提交并推送到 GitHub
    sync_status = None
    if syncer and changed:
        try:
            rel = stock_path.resolve().relative_to(repo_dir)
        except ValueError:
            log("[同步] 库存表不在仓库目录内，跳过同步")
        else:
            msg = (f"库存更新 {datetime.now():%Y-%m-%d %H:%M}: "
                   f"{bom_path.stem} ×{sets} 套，扣减 {len(changed)} 种货品")
            log("[同步] 正在提交并推送库存表 ...")
            try:
                sync_status = syncer.commit_and_push(rel.as_posix(), msg)
                log("[同步] 已推送到 GitHub"
                    if sync_status == "pushed"
                    else "[同步] 库存无变化，未提交")
            except (GitSyncError, subprocess.TimeoutExpired) as e:
                log(f"[同步] 推送失败：{e}")

    return {
        "out_path": out_path,
        "purchase_count": len(data["purchase_rows"]),
        "stats": stats,
        "changed": changed,
        "backup_path": backup_path,
        "sync": sync_status,
    }


# ---------------------------------------------------------------- GUI


def start_gui(selftest: bool = False):
    import os
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    root = tk.Tk()
    root.title("BOM 采购计算器")
    root.geometry("780x620")
    root.minsize(700, 540)

    bom_var = tk.StringVar()
    sets_var = tk.StringVar(value="1")
    stock_var = tk.StringVar(
        value=str(DEFAULT_STOCK) if DEFAULT_STOCK.exists() else "")
    rc_var = tk.BooleanVar(value=DEFAULT_RC.exists())
    upd_var = tk.BooleanVar(value=True)
    bak_var = tk.BooleanVar(value=True)
    git_ok = GitSync(BASE_DIR).available
    sync_var = tk.BooleanVar(value=git_ok)

    # 只有一个 BOM 文件时自动填上
    cands = [p for p in BASE_DIR.glob("*.xlsx")
             if p.name.startswith("BOM") and "比对结果" not in p.name
             and "需采购" not in p.name]
    if len(cands) == 1:
        bom_var.set(str(cands[0]))

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    def add_file_row(row: int, label: str, var: tk.StringVar,
                     on_browse):
        ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w",
                                        padx=(0, 8), pady=5)
        ttk.Entry(frm, textvariable=var).grid(row=row, column=1,
                                              sticky="ew", pady=5)
        ttk.Button(frm, text="浏览...", command=on_browse).grid(
            row=row, column=2, sticky="e", padx=(8, 0), pady=5)

    def pick_bom():
        p = filedialog.askopenfilename(
            title="选择 BOM 文件", initialdir=str(BASE_DIR),
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if p:
            bom_var.set(p)

    def pick_stock():
        p = filedialog.askopenfilename(
            title="选择库存表", initialdir=str(BASE_DIR),
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if p:
            stock_var.set(p)

    add_file_row(0, "BOM 文件：", bom_var, pick_bom)

    ttk.Label(frm, text="生产套数：").grid(row=1, column=0, sticky="w",
                                          padx=(0, 8), pady=5)
    sets_box = ttk.Spinbox(frm, from_=1, to=100000, textvariable=sets_var,
                           width=12)
    sets_box.grid(row=1, column=1, sticky="w", pady=5)
    ttk.Label(frm, text="需求数量 = BOM 数量 × 套数").grid(
        row=1, column=1, sticky="w", padx=(110, 0), pady=5)

    add_file_row(2, "库存表：", stock_var, pick_stock)

    opts = ttk.Frame(frm)
    opts.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 2))
    ttk.Checkbutton(opts, text="使用 0603 阻容免购清单（电阻电容.md）",
                    variable=rc_var).pack(side="left", padx=(0, 16))
    ttk.Checkbutton(opts, text="扣减库存并保存", variable=upd_var).pack(
        side="left", padx=(0, 16))
    ttk.Checkbutton(opts, text="保存前备份库存表", variable=bak_var).pack(
        side="left")

    opts2 = ttk.Frame(frm)
    opts2.grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 2))
    sync_cb = ttk.Checkbutton(
        opts2, text="库存更新后自动同步到 GitHub（多人协作）",
        variable=sync_var)
    sync_cb.pack(side="left")
    if not git_ok:
        sync_cb.config(state="disabled")
        ttk.Label(opts2, text="（当前目录不是 Git 仓库）").pack(
            side="left", padx=(6, 0))

    btns = ttk.Frame(frm)
    btns.grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 4))
    run_btn = ttk.Button(btns, text="开始计算")
    run_btn.pack(side="left")
    open_btn = ttk.Button(btns, text="打开输出文件", state="disabled")
    open_btn.pack(side="left", padx=(8, 0))
    ttk.Button(btns, text="BOM 要求",
               command=lambda: messagebox.showinfo("BOM 表要求", BOM_HELP)
               ).pack(side="left", padx=(8, 0))

    log_text = scrolledtext.ScrolledText(frm, height=18, wrap="none",
                                         font=("Consolas", 9))
    log_text.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
    frm.rowconfigure(6, weight=1)

    last_out = {"path": None}

    def log(msg: str):
        log_text.insert("end", str(msg) + "\n")
        log_text.see("end")
        root.update_idletasks()

    def open_output():
        p = last_out["path"]
        if p and Path(p).exists():
            os.startfile(str(p))  # noqa: S606  (Windows)
        else:
            messagebox.showwarning("提示", "还没有生成输出文件")

    def start():
        bom = bom_var.get().strip()
        stock = stock_var.get().strip()
        if not bom or not Path(bom).exists():
            messagebox.showerror("错误", "请选择有效的 BOM 文件")
            return
        if not stock or not Path(stock).exists():
            messagebox.showerror("错误", "请选择有效的库存表")
            return
        try:
            sets = int(sets_var.get())
        except ValueError:
            messagebox.showerror("错误", "套数必须是整数")
            return
        if sets <= 0:
            messagebox.showerror("错误", "套数必须大于 0")
            return

        log_text.delete("1.0", "end")
        run_btn.config(state="disabled")
        open_btn.config(state="disabled")
        root.config(cursor="watch")
        try:
            res = run(
                bom_path=bom,
                sets=sets,
                stock_path=stock,
                rc_path=DEFAULT_RC if rc_var.get() else None,
                update_stock=upd_var.get(),
                backup=bak_var.get(),
                sync=sync_var.get(),
                repo_dir=BASE_DIR,
                log=log,
            )
            last_out["path"] = res["out_path"]
            open_btn.config(state="normal")
            log("")
            log("完成。")
        except (Exception, SystemExit) as e:
            log(f"出错：{e}")
            messagebox.showerror("出错", str(e))
        finally:
            root.config(cursor="")
            run_btn.config(state="normal")

    run_btn.config(command=start)
    open_btn.config(command=open_output)
    sets_box.bind("<Return>", lambda _e: start())

    if selftest:
        root.after(500, root.destroy)
    root.mainloop()


# ---------------------------------------------------------------- CLI


def cli(argv=None) -> int:
    ap = argparse.ArgumentParser(description="BOM 采购计算器（命令行模式）")
    ap.add_argument("bom", type=Path, help="BOM xlsx")
    ap.add_argument("--sets", type=int, required=True, help="生产套数")
    ap.add_argument("--stock", type=Path, default=None,
                    help="库存 xlsx（默认脚本目录下 库存管理 副本.xlsx）")
    ap.add_argument("-o", "--out", type=Path, default=None,
                    help="输出采购 BOM（默认 <BOM名>_需采购.xlsx）")
    ap.add_argument("--threshold", type=float, default=0.82,
                    help="模糊匹配阈值，默认 0.82")
    ap.add_argument("--no-rc", action="store_true", help="不用 0603 阻容清单")
    ap.add_argument("--dry-run", action="store_true", help="只计算，不扣库存")
    ap.add_argument("--no-backup", action="store_true", help="扣库存前不备份")
    ap.add_argument("--no-sync", action="store_true",
                    help="不同步库存表到 GitHub")
    ap.add_argument("--repo", type=Path, default=None,
                    help="Git 仓库目录（默认脚本 / exe 所在目录）")
    args = ap.parse_args(argv)

    run(
        bom_path=args.bom,
        sets=args.sets,
        stock_path=args.stock,
        out_path=args.out,
        rc_path=None if args.no_rc else DEFAULT_RC,
        threshold=args.threshold,
        update_stock=not args.dry_run,
        backup=not args.no_backup,
        sync=not args.no_sync and not args.dry_run,
        repo_dir=args.repo,
    )
    return 0


def _selftest() -> None:
    """打包后自检：建窗口、定位资源、跑一次比对，结果写入临时文件。"""
    import tempfile

    lines = [f"base={BASE_DIR}"]
    start_gui(selftest=True)
    lines.append("gui=ok")
    lines.append(f"stock={DEFAULT_STOCK.exists()} rc={DEFAULT_RC.exists()} "
                 f"git={GitSync(BASE_DIR).available}")

    cands = sorted(p for p in BASE_DIR.glob("BOM*.xlsx")
                   if "比对结果" not in p.name and "需采购" not in p.name)
    if cands and DEFAULT_STOCK.exists():
        try:
            d = compute(cands[0], DEFAULT_STOCK, 1)
            lines.append(f"compute=ok rows={len(d['results'])} "
                         f"purchase={len(d['purchase_rows'])}")
        except Exception as e:  # noqa: BLE001
            lines.append(f"compute=fail {e}")

    Path(tempfile.gettempdir(), "bom_gui_selftest.txt").write_text(
        "\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        _selftest()
    elif "--cli" in sys.argv:
        sys.exit(cli([a for a in sys.argv[1:] if a != "--cli"]))
    else:
        start_gui()
