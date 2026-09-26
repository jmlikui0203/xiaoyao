# -*- coding: utf-8 -*-
"""
超料记录 资料同步脚本（对应 AGENT.md 硬流程）

执行顺序（AGENT.md 头部约定）:
  §9a 全部补料损耗表 -> loss.duckdb 记录级 upsert (md5 幂等, 已入库跳过/未入库新增;
       **只增不删**: 同名换版/文件移走都保留库内数据;
       目录无损耗表时不清库不报错, 直接用库内现有数据)
  §4  PG视图 production_start_finish 同步 指令/数量/三完成日
  §10 orders 表补全 型号/型体/颜色/交期 (work_no 组合号拆分 + norm 归一化)
  §11 特殊订单(无主指令数量) 分单数量 + 生产累计完成日
  §11.1 特殊订单 型号/型体/颜色/交期 (交期取主指令与分单全局最晚)
  §9  补料损耗表 按截止日汇总 超料1/2/3 + 全量汇总 超料4 (多文件取文件名日期最新)
  §13 超料月统计: 只更新脚本自有列 1~4(按月份键), 新月份整行新增,
       全部月份按月份列升序整行重排(用户列随月份键走, 公式自动平移)
  §5  指令超料 全表格式统一为第 3 行基准(仅修正不一致的单元格)
  §12 按 交期(列6) 从小到大排序(收尾, 已有序则跳过)

幂等/性能约定（2026-09-25 用户要求）:
  - 损耗源文件 md5 与库内 meta 一致 -> 跳过解析与入库
  - 所有写入先与现值比对, 相同不写 (值与样式均差量)
  - 全程无任何改动 -> 不备份、不保存、不改写文件; 确有改动才在保存前备份

运行: E:/tools/uv/uv.exe run python main.py  (在 E:/xiaoyao 下)
"""
import hashlib
import re
import shutil
import glob
import os
import sys
import time
from datetime import datetime, date
from collections import defaultdict
from copy import copy

import openpyxl
import psycopg2
import duckdb
from openpyxl.styles import Font
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

# ---------------- 常量(与 AGENT.md 一致) ----------------
DB = dict(host="localhost", port=5432, dbname="postgres",
          user="postgres", password="123456")

SRC = r"E:/xiaoyao/超料记录.xlsx"
SHEET = "指令超料"
DATE_FMT = "mm-dd-yy"

# 列位置(1-based), AGENT.md §3
COL_INSTR = 1    # 指令
COL_MODEL = 2    # 型号
COL_STYLE = 3    # 型体
COL_COLOR = 4    # 颜色
COL_QTY = 5      # 数量
COL_CFM = 6      # 交期
COL_CUT = 7      # 裁断完成日
COL_CUT_OVER = 8   # 超料1
COL_STITCH = 9   # 针车完成日
COL_ST_OVER = 10   # 超料2
COL_ASM = 11     # 成型完成日
COL_ASM_OVER = 12  # 超料3
COL_SETTLE = 13  # 最终结算日(暂无数据源)
COL_OVER4 = 14   # 超料4 = 该指令全部补料记录金额汇总(§9b)

MARK_NODATA = "-"  # 超料列: 源表查无此指令(无数据), 区别于 0

MONTH_SHEET = "超料月统计"   # §13 月统计表(只维护自有列 1~4)
MONTH_HEADERS = ["月份", "超料总额", "大小码金额", "责任总额"]
YM_PAT = re.compile(r"^\d{4}-\d{2}$")     # 月份键: 2026-07
REASON_BIG_SMALL = "大碼比小碼多"  # 损耗原因: 大小码(不计责任)

# §9 DuckDB 补料损耗库: loss_detail 只存业务列(用户指定 7 列) + 来源文件
DUCKDB_PATH = r"E:/xiaoyao/loss.duckdb"
LOSS_TABLE = "loss_detail"
META_TABLE = "loss_import_meta"   # 入库指纹(file_md5), 用于幂等跳过
HASH_TABLE = "loss_row_hash"      # 已废弃(v3 只增不删不再需要按文件簿记), 仅在自动清理时引用
# 源表列号(1-based) -> 库字段
LOSS_COLMAP = {1: "loss_date", 2: "order_no", 3: "category",
               7: "instr_no", 8: "material_no",
               17: "amount", 20: "reason"}

VIEW_FILTER_FROM = "15D09A001"  # production/视图同口径过滤

norm = lambda s: "".join(str(s).split()).upper()          # 去所有空白+大写
SUB_PAT = re.compile(r"^(.+)-\d{2,}$")                    # 分单: xxx-NN
LOSS_FILE_PAT = re.compile(r"(\d{8})")                    # 文件名中的 8 位日期

# 差量写入计数: 值改动 / 样式改动 (为 0 则不保存不备份)
STATS = {"cells": 0, "styles": 0}


def to_datetime(d):
    """date -> datetime(写入 Excel 用), 其余原样"""
    if isinstance(d, datetime):
        return d
    if isinstance(d, date):
        return datetime(d.year, d.month, d.day)
    return d


def pdate(v):
    """宽容日期解析: datetime/date/多种字符串 -> date 或 None"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip()
        for f in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, f).date()
            except ValueError:
                pass
    return None


def values_equal(a, b):
    """单元格现值与目标值是否等价(空串=None, 数值容差, 日期同值)"""
    if isinstance(a, str) and a.strip() == "":
        a = None
    if isinstance(b, str) and b.strip() == "":
        b = None
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 1e-6
    if isinstance(a, (datetime, date)) and isinstance(b, (datetime, date)):
        return pdate(a) == pdate(b)
    return a == b


def set_cell(ws, r, c, value, number_format=None):
    """差量写入: 与现值相同则不写; 返回是否实际改动"""
    cell = ws.cell(r, c)
    if values_equal(cell.value, value):
        return False
    cell.value = value
    if number_format:
        cell.number_format = number_format
    STATS["cells"] += 1
    return True


def _color_sig(c):
    if c is None:
        return None
    return (getattr(c, "type", None), getattr(c, "rgb", None),
            getattr(c, "theme", None), getattr(c, "tint", None),
            getattr(c, "indexed", None))


def _sig_parts(f, fi, b, a, numfmt, p):
    """样式六要素 -> 签名(供 style_sig 与整行重排的目标样式比对复用)"""
    return (
        (f.name, f.sz, f.b, f.i, f.u, f.strike, _color_sig(f.color)),
        (fi.patternType, _color_sig(fi.fgColor), _color_sig(fi.bgColor)),
        tuple((getattr(b, s).style, _color_sig(getattr(b, s).color))
              for s in ("left", "right", "top", "bottom", "diagonal")),
        (a.horizontal, a.vertical, a.wrapText, a.indent),
        numfmt,
        (p.locked, p.hidden),
    )


def style_sig(cell):
    """单元格样式签名(用于差量样式比对)"""
    return _sig_parts(cell.font, cell.fill, cell.border, cell.alignment,
                      cell.number_format, cell.protection)


def backup_src():
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    bak = rf"E:/xiaoyao/超料记录_备份_{ts}.xlsx"
    shutil.copy(SRC, bak)
    print(f"[备份] {bak}")
    return bak


def list_loss_files():
    """§9: 目录下全部补料损耗表(文件名末尾 8 位日期)都是数据源, 按文件合并入库
    (2026-09-26 用户规则: 不同期间的文件=追加记录, 不再"只取最新一份")
    目录一份都没有 -> 不报错, 返回空列表, §9a 改为直接使用库内现有数据
    (2026-09-26 用户规则: 没有损耗表时直接用 duckdb 跑)
    同一天存在多份 -> 报错停下, 请用户人工确认(防重复算钱)"""
    files = []
    for pat in (r"E:/xiaoyao/補料發料損耗表-*.xlsx", r"E:/xiaoyao/补料发料损耗表-*.xlsx"):
        files += glob.glob(pat)
    if not files:
        print("[§9] 目录下没有补料损耗表, 将直接使用 loss.duckdb 库内现有数据")
        return []
    dated = []
    for f in files:
        m = LOSS_FILE_PAT.search(os.path.basename(f))
        if not m:
            raise ValueError(f"补料损耗表文件名中提取不到 8 位日期: {f}")
        dated.append((m.group(1), f))
    seen = {}
    for d, f in dated:
        if d in seen:
            raise ValueError(f"同一天存在多份补料损耗表, 需人工确认: {[seen[d], f]}")
        seen[d] = f
    dated.sort()
    print(f"[§9] 补料损耗表共 {len(dated)} 份, 全部并入: "
          + ", ".join(os.path.basename(f) for _, f in dated))
    return dated


def file_md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def fetch_pg_data():
    """拉取 PG: 视图 / orders / order_sizerun / production"""
    conn = psycopg2.connect(connect_timeout=10, **DB)
    cur = conn.cursor()

    # §2 视图
    cur.execute("""SELECT work_no, order_qty, cut_finish_date,
                          stitch_finish_date, assembly_finish_date
                   FROM production_start_finish""")
    view = []
    for wno, qty, cf, sf, af in cur.fetchall():
        if wno:
            view.append((str(wno).strip(), qty, cf, sf, af))

    # §10 orders
    cur.execute("SELECT work_no, style_no, style_name, color, current_cfm FROM orders")
    orders = cur.fetchall()

    # §11 order_sizerun (分单数量 / 裸键数量)
    cur.execute("SELECT work_no, quantity FROM order_sizerun")
    osr_rows = cur.fetchall()

    # §11 production (与视图同口径过滤)
    cur.execute(f"""SELECT work_no, work_date, cutting, stitching, assembling
                    FROM production WHERE work_no > '{VIEW_FILTER_FROM}'
                    ORDER BY work_no, work_date""")
    prod_rows = cur.fetchall()

    conn.close()

    # orders: 组合 work_no 按 / 拆分 -> 分量映射(冲突取最晚 cfm 由 §11.1 逻辑处理)
    comp_map = defaultdict(list)
    for w, sno, sname, col, cfm in orders:
        if not w:
            continue
        for c in str(w).replace(" ", "").split("/"):
            c = c.strip()
            if c:
                comp_map[norm(c)].append(
                    dict(sno=sno, sname=sname, col=col, cfm=cfm, src=str(w)))

    # order_sizerun: 分单分量 -> 数量和; 裸键 -> 有效数量和
    sub_qty = defaultdict(int)
    bare_qty = defaultdict(int)
    for w, q in osr_rows:
        if not w or q is None:
            continue
        for c in str(w).replace(" ", "").split("/"):
            c = c.strip()
            if not c:
                continue
            if SUB_PAT.match(c):
                sub_qty[norm(c)] += q
            else:
                bare_qty[norm(c)] += q

    # production: 主指令 -> 按日期的生产序列
    prod = defaultdict(list)
    for w, d, cut, st, asm in prod_rows:
        if w and d:
            prod[str(w).strip().upper()].append(
                (d, cut or 0, st or 0, asm or 0))

    print(f"[PG] 视图 {len(view)} 行 | orders {len(orders)} 行 | "
          f"order_sizerun {len(osr_rows)} 行 | production {len(prod_rows)} 行")
    return view, comp_map, sub_qty, bare_qty, prod


def step4_sync_view(ws, view):
    """§4: 视图同步(命中更新/未命中追加); NULL 不写入不清空; 差量写入"""
    existing = {}
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, COL_INSTR).value
        if v is not None:
            existing[norm(v)] = r

    updated = appended = 0
    next_row = (max(existing.values()) if existing else 1) + 1
    for wno, qty, cf, sf, af in view:
        key = norm(wno)
        vals = ((COL_QTY, qty), (COL_CUT, cf), (COL_STITCH, sf), (COL_ASM, af))
        if key in existing:
            r = existing[key]
            updated += 1
        else:
            r = next_row
            next_row += 1
            set_cell(ws, r, COL_INSTR, wno)
            appended += 1
        for col, val in vals:
            if val is not None:  # NULL 不写入不清空
                set_cell(ws, r, col,
                         to_datetime(val) if col != COL_QTY else val,
                         DATE_FMT if col != COL_QTY else None)
    print(f"[§4] 视图同步: 更新 {updated} 行, 追加 {appended} 行")


def special_subs(m, sub_qty, bare_qty):
    """§11 判定特殊订单: 主指令无有效数量且存在 {m}-NN 分单 -> 返回分单 dict, 否则 None
    (判定不能用"裸键不存在", 裸键可能 quantity=NULL, 必须看有效数量合计)"""
    if bare_qty.get(m, 0) > 0:
        return None
    subs = {k: q for k, q in sub_qty.items()
            if k.startswith(m + "-") and SUB_PAT.match(k)}
    return subs or None


def step10_orders_fill(ws, comp_map, sub_qty, bare_qty):
    """§10: 型号/型体/颜色/交期 (组合号已拆分, norm 匹配; 交期冲突默认取最晚)

    特殊订单行(§11 判定命中)在此跳过: 其 款式/交期 统一由 §11.1 填写
    (交期=主指令与分单全局最晚)。否则 §10 先写主指令交期、§11.1 再改最晚,
    每次运行都会产生无意义的反复重写(2026-09-25 修正)。
    """
    filled = unmatched = n_special = 0
    cfm_conflicts = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, COL_INSTR).value
        if v is None:
            continue
        m = norm(v)
        if special_subs(m, sub_qty, bare_qty):
            n_special += 1
            continue
        recs = comp_map.get(m)
        if not recs:
            unmatched += 1
            continue
        d = recs[0]
        set_cell(ws, r, COL_MODEL, d["sno"])
        set_cell(ws, r, COL_STYLE, d["sname"])
        set_cell(ws, r, COL_COLOR, d["col"])
        cfms = [x["cfm"] for x in recs if x["cfm"] is not None]
        if len({str(x) for x in cfms}) > 1:
            cfm_conflicts.append((str(v), max(cfms)))  # 冲突默认取最晚(AGENT.md §10)
        if cfms:
            set_cell(ws, r, COL_CFM, to_datetime(max(cfms)), DATE_FMT)
        filled += 1
    print(f"[§10] orders 四列: 填充 {filled} 行, 未匹配 {unmatched} 行, "
          f"特殊订单转§11 {n_special} 行"
          + (f", 交期冲突 {len(cfm_conflicts)} 处(已取最晚: {cfm_conflicts})" if cfm_conflicts else ""))


def finish_dates_per_proc(rows, qty):
    """§11: 各工序独立判定 生产累计>=qty 的最早日期 -> (裁断,针车,成型)"""
    out = [None, None, None]
    acc = [0, 0, 0]
    for d, c, s, a in rows:
        acc[0] += c; acc[1] += s; acc[2] += a
        for i in range(3):
            if out[i] is None and qty and acc[i] >= qty:
                out[i] = d
    return out


def step11_special_orders(ws, comp_map, sub_qty, bare_qty, prod):
    """§11+§11.1: 特殊订单 分单数量/完成日/款式/交期(全局最晚)"""
    n = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, COL_INSTR).value
        if v is None:
            continue
        m = norm(v)
        subs = special_subs(m, sub_qty, bare_qty)
        if not subs:
            continue
        n += 1
        # 数量 = 分单求和
        total = sum(subs.values())
        if total > 0:
            set_cell(ws, r, COL_QTY, total)
        # 完成日 = 主指令生产累计 >= 分单合计 的最早日期(工序独立)
        cf, sf, af = finish_dates_per_proc(prod.get(m, []), total)
        for col, fd in ((COL_CUT, cf), (COL_STITCH, sf), (COL_ASM, af)):
            if fd is not None:
                set_cell(ws, r, col, to_datetime(fd), DATE_FMT)
        # §11.1 款式: 主指令直查优先, 否则分单唯一款式
        direct = comp_map.get(m)
        if direct:
            sno, sname, col_ = direct[0]["sno"], direct[0]["sname"], direct[0]["col"]
        else:
            styles = {(x["sno"], x["sname"], x["col"])
                      for lst in [comp_map[k] for k in subs] for x in lst}
            if len(styles) == 1:
                sno, sname, col_ = styles.pop()
            else:
                print(f"[§11.1] 警告: {v} 分单款式不唯一, 跳过款式填充: {styles}")
                sno = sname = col_ = None
        if sno is not None:
            set_cell(ws, r, COL_MODEL, sno)
            set_cell(ws, r, COL_STYLE, sname)
            set_cell(ws, r, COL_COLOR, col_)
        # §11.1 交期 = max(主指令, 全部分单)
        all_cfm = [x["cfm"] for k in subs for x in comp_map[k] if x["cfm"] is not None]
        if direct and direct[0]["cfm"] is not None:
            all_cfm.append(direct[0]["cfm"])
        if all_cfm:
            set_cell(ws, r, COL_CFM, to_datetime(max(all_cfm)), DATE_FMT)
    print(f"[§11] 特殊订单(分单汇总): 处理 {n} 行")


def parse_loss_file(loss_file):
    """解析一份补料损耗表 -> (rows, months, skipped); 只取用户指定 7 列 + 来源文件名"""
    lws = openpyxl.load_workbook(loss_file, data_only=True).worksheets[0]
    src_name = os.path.basename(loss_file)
    rows = []
    months = set()
    skipped = 0
    for r in range(2, lws.max_row + 1):
        ino = lws.cell(r, 7).value
        if ino is None:
            skipped += 1
            continue
        amt = lws.cell(r, 17).value
        try:
            amt = float(amt) if amt is not None else None
        except (TypeError, ValueError):
            amt = None
        vals = {
            1: pdate(lws.cell(r, 1).value),
            2: str(lws.cell(r, 2).value).strip() if lws.cell(r, 2).value is not None else None,
            3: str(lws.cell(r, 3).value).strip() if lws.cell(r, 3).value is not None else None,
            7: str(ino).strip(),
            8: str(lws.cell(r, 8).value).strip() if lws.cell(r, 8).value is not None else None,
            17: amt,
            20: str(lws.cell(r, 20).value).strip() if lws.cell(r, 20).value is not None else None,
        }
        if vals[1] is not None:
            months.add(vals[1].strftime("%Y-%m"))
        rows.append(tuple(vals[c] for c in sorted(LOSS_COLMAP)) + (src_name,))
    return rows, months, skipped


def _row_hash(biz7):
    """行内容指纹: 7 个业务字段的 repr md5。
    源表无天然业务主键, 且实测存在 7 字段全同的"真重复"行(同單號同材料同金额
    的真实多笔补料, 20260925 份 95 组 / 20260926 份 28 组), 故 upsert 不能只按
    hash 去重到 1 份, 必须按份数处理(见 import_loss_to_duckdb)。"""
    return hashlib.md5(repr(tuple(biz7)).encode("utf-8")).hexdigest()


def import_loss_to_duckdb(loss_files):
    """§9a: 目录下全部补料损耗表 -> loss.duckdb 记录级 upsert(只增不删)
    (2026-09-26 用户规则 v3: "每一次损耗表的更新都是 upsert——判断哪些已入库、
    哪些未入库, 未入库的才新增"; 已入库的一律保留, 上半年数据不能删)

    - 每行以 7 个业务字段的 hash 标识: 库内已有该 hash -> 跳过;
      没有 -> 插入(文件内同 hash 多份是真实多笔, 首次入库按文件内份数全量插入)
    - 同名文件换版 -> 只新增库里没有的行, **旧行一律保留, 不替换不删**
    - 文件移出目录 -> 只提示, **数据保留**
    - 月份重叠 -> 仅警告(upsert 按行内容天然防双算), 不停止
    - 文件 md5 与 loss_import_meta 一致 -> 整文件跳过解析(幂等)
    注意: 只增不删意味着"源文件修正某行"在库里表现为新增修正行、旧行仍保留;
    确需删除/修正库内数据须人工明确提出后单独处理, 不能靠覆盖文件实现。
    loss_detail 只含 7 业务列 + source_file(用户指定)。
    """
    if not loss_files:
        # 目录无损耗表(2026-09-26 用户规则): 不报错、不清库, 直接使用库内现有数据。
        # 注意必须提前返回——否则下方"文件移走则删除"逻辑会把所有行误判为已移除。
        con = duckdb.connect(DUCKDB_PATH)
        try:
            tables = {r[0] for r in con.execute(
                "SELECT table_name FROM information_schema.tables").fetchall()}
            n = (con.execute(f"SELECT COUNT(*) FROM {LOSS_TABLE}").fetchone()[0]
                 if LOSS_TABLE in tables else 0)
        finally:
            con.close()
        if n == 0:
            sys.exit("[§9a] 目录无损耗表且 loss.duckdb 内也无数据, "
                     "无法计算超料, 请先放入补料损耗表")
        print(f"[§9a] 无损耗表文件, 跳过入库, 直接使用库内现有 {n} 行")
        return n
    con = duckdb.connect(DUCKDB_PATH)
    try:
        tables = {r[0] for r in con.execute(
            "SELECT table_name FROM information_schema.tables").fetchall()}
        has_detail = LOSS_TABLE in tables
        meta = {}
        if META_TABLE in tables:
            for sf, md5v in con.execute(
                    f"SELECT source_file, file_md5 FROM {META_TABLE}").fetchall():
                meta[sf] = md5v

        con.execute(f"""CREATE TABLE IF NOT EXISTS {LOSS_TABLE} (
            loss_date    DATE,
            order_no     VARCHAR,
            category     VARCHAR,
            instr_no     VARCHAR,
            material_no  VARCHAR,
            amount       DOUBLE,
            reason       VARCHAR,
            source_file  VARCHAR)""")
        con.execute(f"""CREATE TABLE IF NOT EXISTS {META_TABLE} (
            file_md5    VARCHAR,
            source_file VARCHAR,
            imported_at TIMESTAMP,
            row_count   BIGINT)""")
        if HASH_TABLE in tables:   # v3 起只增不删, 不再需要按文件簿记
            con.execute(f"DROP TABLE {HASH_TABLE}")

        # 阶段1: 解析有变化的文件(md5 一致 -> 整文件跳过)
        parsed = {}    # name -> (md5, rows, months, skipped)
        for _, path in loss_files:
            name = os.path.basename(path)
            md5 = file_md5(path)
            if has_detail and meta.get(name) == md5:
                print(f"[§9a] {name} 未变化(md5 一致), 跳过")
                continue
            rows, months, skipped = parse_loss_file(path)
            parsed[name] = (md5, rows, months, skipped)

        # 月份重叠提示(upsert 按行内容防双算, 仅警告)
        present = {os.path.basename(f) for _, f in loss_files}
        month_files = defaultdict(set)
        for name, (_, _, months, _) in parsed.items():
            for ym in months:
                month_files[ym].add(name)
        if has_detail:
            for sf, ym in con.execute(
                    f"""SELECT DISTINCT source_file, strftime(loss_date, '%Y-%m')
                        FROM {LOSS_TABLE} WHERE loss_date IS NOT NULL""").fetchall():
                if sf in present and sf not in parsed:
                    month_files[ym].add(sf)
        overlaps = {ym: sorted(s) for ym, s in month_files.items() if len(s) > 1}
        if overlaps:
            print(f"[§9a] 提示: 不同文件月份覆盖重叠 {overlaps}; "
                  f"相同内容的行不会重复入库, 仅提示人工留意")

        # 库内现有行 hash 集合(只判在/不在; 2 万行 python 端计算, 毫秒级)
        existing = set()
        for row in con.execute(
                f"""SELECT loss_date, order_no, category,
                           instr_no, material_no, amount, reason
                    FROM {LOSS_TABLE}""").fetchall():
            existing.add(_row_hash(row))

        # 阶段2: 事务内只增不删 upsert
        con.execute("BEGIN TRANSACTION")
        inserted = 0
        for name, (md5, rows, months, skipped) in parsed.items():
            cnt = defaultdict(int)
            payload = {}
            for t in rows:
                h = _row_hash(t[:7])
                cnt[h] += 1
                payload.setdefault(h, t[:7])
            n_new = 0
            for h, n in cnt.items():
                if h in existing:                    # 已入库 -> 跳过
                    continue
                con.executemany(
                    f"INSERT INTO {LOSS_TABLE} VALUES (?,?,?,?,?,?,?,?)",
                    [tuple(payload[h]) + (name,)] * n)
                existing.add(h)
                n_new += n
            inserted += n_new
            con.execute(f"DELETE FROM {META_TABLE} WHERE source_file = ?", [name])
            con.execute(f"INSERT INTO {META_TABLE} VALUES (?, ?, now(), ?)",
                        [md5, name, len(rows)])
            print(f"[§9a] 解析 {name}: {len(rows)} 行 -> 新增 {n_new} 行 / "
                  f"已入库跳过 {len(rows) - n_new} 行"
                  + (f", 无指令号跳过 {skipped} 行" if skipped else ""))
        con.execute("COMMIT")

        # 文件移出目录: 只提示, 数据一律保留(只增不删, 2026-09-26 用户规则 v3)
        for old in sorted(set(meta) - present):
            print(f"[§9a] 提示: {old} 已不在目录, 其数据仍保留在库中(只增不删)")

        print(f"[§9a] upsert: 新增 {inserted} 行 / 删除 0 行(规则: 只增不删)")
        n_total = con.execute(f"SELECT COUNT(*) FROM {LOSS_TABLE}").fetchone()[0]
        n_months = con.execute(
            f"""SELECT COUNT(DISTINCT strftime(loss_date, '%Y-%m'))
                FROM {LOSS_TABLE} WHERE loss_date IS NOT NULL""").fetchone()[0]
        print(f"[§9a] 库内合计 {n_total} 行 / {n_months} 个月份")
        return n_total
    finally:
        con.close()


def step9_loss_summary(ws):
    """§9: 补料责任金额按三截止日累加 -> 超料1/2/3 (含 -/0/空 三态语义)
       §9b: 超料4 = 该指令全部补料记录金额合计(不限截止日)

    数据从 loss.duckdb 读取(§9a 已入库); 参与汇总的行口径与旧版一致:
    指令號非空 + 補料日期可解析 + 補料責任金額为数值。差量写入。
    """
    con = duckdb.connect(DUCKDB_PATH, read_only=True)
    try:
        recs = con.execute(f"""SELECT instr_no, loss_date, amount
                               FROM {LOSS_TABLE}
                               WHERE loss_date IS NOT NULL
                                 AND amount IS NOT NULL""").fetchall()
    finally:
        con.close()
    loss = defaultdict(list)
    for ino, d, amt in recs:
        loss[norm(ino)].append((d, amt))

    n_num = n_mark = n_blank = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, COL_INSTR).value
        if v is None:
            continue
        rows = loss.get(norm(v))
        for col_d, col_a in ((COL_CUT, COL_CUT_OVER),
                             (COL_STITCH, COL_ST_OVER),
                             (COL_ASM, COL_ASM_OVER)):
            cutoff = pdate(ws.cell(r, col_d).value)
            if cutoff is None:                       # 完成日空 -> 留空
                set_cell(ws, r, col_a, None)
                n_blank += 1
            elif rows:                               # 源表内 -> 累计和(含0)
                set_cell(ws, r, col_a,
                         round(sum(a for d, a in rows if d <= cutoff), 2))
                n_num += 1
            else:                                    # 源表查无 -> 标记
                set_cell(ws, r, col_a, MARK_NODATA)
                n_mark += 1
    print(f"[§9] 超料1/2/3 汇总: 数值 {n_num} | 标记'-' {n_mark} | 完成日空留空 {n_blank}")

    # §9b: 超料4 = 该指令全部补料记录金额合计(不限截止日, NULL 金额不计入)
    con = duckdb.connect(DUCKDB_PATH, read_only=True)
    try:
        total_recs = con.execute(
            f"SELECT instr_no, amount FROM {LOSS_TABLE}").fetchall()
    finally:
        con.close()
    present = set()
    total = defaultdict(float)
    for ino, amt in total_recs:
        k = norm(ino)
        present.add(k)
        if amt is not None:
            total[k] += amt

    n4_num = n4_mark = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, COL_INSTR).value
        if v is None:
            continue
        k = norm(v)
        if k in present:                             # 库内有记录 -> 全量合计(含0)
            set_cell(ws, r, COL_OVER4, round(total[k], 2))
            n4_num += 1
        else:                                        # 库内查无 -> 标记
            set_cell(ws, r, COL_OVER4, MARK_NODATA)
            n4_mark += 1
    print(f"[§9b] 超料4 全量汇总: 数值 {n4_num} | 标记'-' {n4_mark}")


def _style_month_row(wb, ws, r):
    """月统计新行样式: 月份列=「指令超料」列1基准, 金额列=列8基准 (与原建表一致)"""
    src = wb[SHEET]
    for c in range(1, 5):
        b = src.cell(3, 1) if c == 1 else src.cell(3, 8)
        cell = ws.cell(r, c)
        cell.font = copy(b.font); cell.fill = copy(b.fill)
        cell.border = copy(b.border); cell.alignment = copy(b.alignment)
        cell.number_format = b.number_format
        cell.protection = copy(b.protection)
    h = src.row_dimensions[3].height
    if h is not None:
        ws.row_dimensions[r].height = h


def _finish_month_sheet(ws):
    ws.column_dimensions["A"].width = 12
    for col in ("B", "C", "D"):
        ws.column_dimensions[col].width = 14
    ws.freeze_panes = "A2"


def _read_row_rec(ws, r, max_col):
    """把整行读成可携带的记录: 值 + 样式六要素副本 + 行高 + 原行号(公式平移基准)"""
    cells = {}
    for c in range(1, max_col + 1):
        cell = ws.cell(r, c)
        cells[c] = (cell.value, copy(cell.font), copy(cell.fill), copy(cell.border),
                    copy(cell.alignment), cell.number_format, copy(cell.protection))
    return {"cells": cells, "height": ws.row_dimensions[r].height, "src_row": r}


def _new_month_rec(wb, max_col):
    """新月份整行记录: 脚本列 1~4 给基准样式(同 _style_month_row), 用户列留空不设定样式"""
    src = wb[SHEET]
    cells = {}
    for c in range(1, max_col + 1):
        if c <= 4:
            b = src.cell(3, 1) if c == 1 else src.cell(3, 8)
            cells[c] = (None, copy(b.font), copy(b.fill), copy(b.border),
                        copy(b.alignment), b.number_format, copy(b.protection))
        else:
            cells[c] = (None, None, None, None, None, None, None)  # 样式 None=写回时不触碰
    return {"cells": cells, "height": src.row_dimensions[3].height, "src_row": None}


def _write_row_rec(ws, r, rec, max_col):
    """把记录写回目标行: 值差量(set_cell), 样式差量(style_sig 比对),
    公式字符串随 (原行->目标行) 平移(Translator, 相对引用自动调整)"""
    for c in range(1, max_col + 1):
        value, font, fill, border, align, numfmt, prot = rec["cells"][c]
        if (isinstance(value, str) and value.startswith("=")
                and rec["src_row"] is not None and rec["src_row"] != r):
            try:
                value = Translator(
                    value, origin=f"{get_column_letter(c)}{rec['src_row']}"
                ).translate_formula(f"{get_column_letter(c)}{r}")
            except Exception:
                pass            # 平移失败则保留原公式, 不阻断流程
        cell = ws.cell(r, c)
        set_cell(ws, r, c, value)
        if font is None:        # 用户列留空位: 不触碰样式
            continue
        if style_sig(cell) != _sig_parts(font, fill, border, align, numfmt, prot):
            cell.font = copy(font); cell.fill = copy(fill)
            cell.border = copy(border); cell.alignment = copy(align)
            cell.number_format = numfmt
            cell.protection = copy(prot)
            STATS["styles"] += 1
    if rec["height"] is not None and ws.row_dimensions[r].height != rec["height"]:
        ws.row_dimensions[r].height = rec["height"]
        STATS["styles"] += 1


def step13_month_summary(wb):
    """§13: 超料月统计 <- loss.duckdb 按月汇总

    只维护脚本自有列 1~4(月份/超料总额/大小码金额/责任总额), 以月份为键整行管理:
      - 已有月份行: 仅当列 2~4 值变化才覆盖(差量)
      - 新月份: **整行新增**(2026-09-26 用户要求), 用户列随月份键走、不错位
      - 新增后**全部月份按月份列从小到大整行重排**(用户要求); 行内公式随
        行号平移(Translator), 相对引用自动指向新行
      - 库内已无的旧月份: 整行保留不删, 只警告
      - 用户自建列(第 5 列起)与任何脚本外内容: 一律不动、不删
      - 表头第 1~4 列与定义不符(被插列/改名) -> 跳过本表并警告, 不盲目覆盖
    列口径:
      超料总额   = 当月全部补料责任金额合计 (loss_date/amount 均非空的行)
      大小码金额 = 当月 損耗原因='大碼比小碼多' 的金额合计 (用户指定口径)
      责任总额   = 超料总额 - 大小码金额
    """
    con = duckdb.connect(DUCKDB_PATH, read_only=True)
    try:
        rows = con.execute(f"""
            SELECT strftime(loss_date, '%Y-%m')                AS ym,
                   ROUND(SUM(amount), 2)                       AS total,
                   ROUND(SUM(CASE WHEN reason = ?
                                  THEN amount ELSE 0 END), 2)  AS big_small
            FROM {LOSS_TABLE}
            WHERE loss_date IS NOT NULL AND amount IS NOT NULL
            GROUP BY ym ORDER BY ym
        """, [REASON_BIG_SMALL]).fetchall()
    finally:
        con.close()
    computed = {ym: (total, big_small, round(total - big_small, 2))
                for ym, total, big_small in rows}

    if MONTH_SHEET not in wb.sheetnames:   # 首次建表
        ws = wb.create_sheet(MONTH_SHEET)
        src = wb[SHEET]
        ws.append(MONTH_HEADERS)
        for c in range(1, 5):
            cell = ws.cell(1, c)
            cell.font = Font(bold=True)
            h = src.cell(1, c)
            cell.border = copy(h.border)
            cell.alignment = copy(h.alignment)
            cell.fill = copy(h.fill)
        for i, ym in enumerate(sorted(computed), start=2):
            total, big_small, resp = computed[ym]
            for c, v in enumerate((ym, total, big_small, resp), start=1):
                ws.cell(i, c).value = v
            _style_month_row(wb, ws, i)
        _finish_month_sheet(ws)
        STATS["cells"] += 4 + 4 * len(computed)
        print(f"[§13] 超料月统计: 新建 {len(computed)} 个月份 "
              f"(超料总额 {sum(v[0] for v in computed.values()):,.2f} / "
              f"大小码 {sum(v[1] for v in computed.values()):,.2f} / "
              f"责任 {sum(v[0] - v[1] for v in computed.values()):,.2f})")
        return

    ws = wb[MONTH_SHEET]
    headers = [ws.cell(1, c).value for c in range(1, 5)]
    headers = [str(h).strip() if isinstance(h, str) else h for h in headers]
    if headers != MONTH_HEADERS:
        print(f"[§13] 警告: 超料月统计 表头 {headers} 与脚本定义 {MONTH_HEADERS} 不符"
              f"(可能被插列/改名), 已跳过本表更新, 请人工核对后恢复表头")
        return

    # ---- 整行记录化(2026-09-26 用户规则): 以月份为键携带整行(含用户列) ----
    max_col = max(4, ws.max_column)
    month_rec = {}     # ym -> rec
    other_recs = []    # 非月份行(用户备注等), 保持相对顺序排在月份块之后
    interspersed = False
    for r in range(2, ws.max_row + 1):
        rec = _read_row_rec(ws, r, max_col)
        v = ws.cell(r, 1).value
        if isinstance(v, str) and YM_PAT.match(v.strip()):
            if other_recs:
                interspersed = True
            month_rec[v.strip()] = rec
        elif any(rec["cells"][c][0] not in (None, "") for c in range(1, max_col + 1)):
            other_recs.append(rec)
        # 全空行直接丢弃(布局重排会自然消除空洞)

    # 新月份: 整行新增(脚本列给基准样式, 用户列留空)
    appended = 0
    for ym in sorted(set(computed) - set(month_rec)):
        rec = _new_month_rec(wb, max_col)
        total, big_small, resp = computed[ym]
        for c, v in ((1, ym), (2, total), (3, big_small), (4, resp)):
            rec["cells"][c] = (v,) + rec["cells"][c][1:]
        month_rec[ym] = rec
        appended += 1

    # 已有月份: 刷新脚本自有列 2~4 的目标值(差量在写回时判定)
    for ym, (total, big_small, resp) in computed.items():
        rec = month_rec[ym]
        for c, v in ((2, total), (3, big_small), (4, resp)):
            rec["cells"][c] = (v,) + rec["cells"][c][1:]

    stale = sorted(set(month_rec) - set(computed))
    if stale:
        print(f"[§13] 警告: 旧月份 {stale} 库内已无, 整行保留不删, 请人工确认")
    if interspersed:
        print("[§13] 警告: 月份块中夹杂非月份行, 已将其移到月份块之后")

    # ---- 按月份升序整行重排(用户列随月份键整体移动, 不错位) ----
    layout = [month_rec[ym] for ym in sorted(month_rec)] + other_recs
    moved = sum(1 for i, rec in enumerate(layout, start=2) if rec["src_row"] != i)
    for i, rec in enumerate(layout, start=2):
        _write_row_rec(ws, i, rec, max_col)
    # 清理布局缩短后的残留行(当前只会因丢弃空行而缩短)
    for r in range(2 + len(layout), ws.max_row + 1):
        for c in range(1, max_col + 1):
            set_cell(ws, r, c, None)

    tot = sum(v[0] for v in computed.values())
    bs = sum(v[1] for v in computed.values())
    print(f"[§13] 超料月统计: 整行新增 {appended} 行 / 按月份升序重排 {moved} 行 "
          f"(共 {len(month_rec)} 个月份; 超料总额 {tot:,.2f} / 大小码 {bs:,.2f} / "
          f"责任 {tot - bs:,.2f})")


def step5_unify_style(ws):
    """§5: 全部数据行统一为第 3 行原始样式(用户否决第 2 行); 差量: 仅修正不一致单元格"""
    base_cells = {}
    for c in range(1, ws.max_column + 1):
        base_cells[c] = ws.cell(3, c)
    base_sigs = {c: style_sig(cell) for c, cell in base_cells.items()}
    height = ws.row_dimensions[3].height
    n = 0
    for r in range(2, ws.max_row + 1):
        if height is not None and ws.row_dimensions[r].height != height:
            ws.row_dimensions[r].height = height
            STATS["styles"] += 1
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if style_sig(cell) != base_sigs[c]:
                b = base_cells[c]
                cell.font = copy(b.font); cell.fill = copy(b.fill)
                cell.border = copy(b.border); cell.alignment = copy(b.alignment)
                cell.number_format = b.number_format
                cell.protection = copy(b.protection)
                n += 1
    STATS["styles"] += n
    print(f"[§5] 格式统一: 第 3 行基准, 修正不一致 {n} 格 (共 {ws.max_row - 1} 行)")


def step12_sort_by_cfm(ws):
    """§12: 按交期升序(整行移动), 空值排末尾(稳定排序); 已有序则跳过"""
    maxc = ws.max_column
    rows = []
    for r in range(2, ws.max_row + 1):
        rows.append(tuple(ws.cell(r, c).value for c in range(1, maxc + 1)))

    def key(vals):
        v = vals[COL_CFM - 1]
        if isinstance(v, datetime):
            return (0, v)
        if isinstance(v, date):
            return (0, datetime(v.year, v.month, v.day))
        return (1, datetime.min)

    keys = [key(v) for v in rows]
    if all(keys[i] >= keys[i - 1] for i in range(1, len(keys))):
        print(f"[§12] 交期排序: 已有序, 跳过 ({len(rows)} 行)")
        return
    rows_sorted = sorted(rows, key=key)   # list.sort 稳定
    for i, vals in enumerate(rows_sorted, start=2):
        for c, v in enumerate(vals, start=1):
            set_cell(ws, i, c, v)
    n_with = sum(1 for vals in rows_sorted if key(vals)[0] == 0)
    print(f"[§12] 交期排序: 重排 {len(rows_sorted)} 行 (有交期 {n_with}, 空排末尾 {len(rows_sorted) - n_with})")


def verify(ws):
    """§6/§9/§12 校验"""
    n_rows = sum(1 for r in range(2, ws.max_row + 1)
                 if ws.cell(r, COL_INSTR).value is not None)
    cnt = {name: 0 for name in ("数量", "交期", "裁断", "针车", "成型")}
    mono_viol = sort_viol = 0
    prev_cfm = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, COL_INSTR).value is None:
            continue
        if ws.cell(r, COL_QTY).value not in (None, ""):
            cnt["数量"] += 1
        cfm = pdate(ws.cell(r, COL_CFM).value)
        if cfm:
            cnt["交期"] += 1
            if prev_cfm is not None and cfm < prev_cfm:
                sort_viol += 1
            prev_cfm = cfm
        for name, col in (("裁断", COL_CUT), ("针车", COL_STITCH), ("成型", COL_ASM)):
            if pdate(ws.cell(r, col).value):
                cnt[name] += 1
        nums = [ws.cell(r, c).value for c in (COL_CUT_OVER, COL_ST_OVER, COL_ASM_OVER)]
        nums = [x for x in nums if isinstance(x, (int, float))]
        if len(nums) == 3 and not (nums[0] <= nums[1] <= nums[2]):
            mono_viol += 1
    print(f"[校验] 数据行 {n_rows} | 已填: {cnt} | "
          f"超料单调性违规 {mono_viol} | 交期升序违规 {sort_viol}")
    return n_rows, sort_viol, mono_viol


def main():
    t0 = time.time()
    if not os.path.exists(SRC):
        sys.exit(f"目标文件不存在: {SRC}")
    loss_files = list_loss_files()
    import_loss_to_duckdb(loss_files)   # §9a: 记录级 upsert 入库(只增不删, md5 幂等)
    view, comp_map, sub_qty, bare_qty, prod = fetch_pg_data()

    wb = openpyxl.load_workbook(SRC)
    ws = wb[SHEET]

    step4_sync_view(ws, view)
    step10_orders_fill(ws, comp_map, sub_qty, bare_qty)
    step11_special_orders(ws, comp_map, sub_qty, bare_qty, prod)
    step9_loss_summary(ws)
    step13_month_summary(wb)
    step5_unify_style(ws)
    step12_sort_by_cfm(ws)

    if STATS["cells"] > 0 or STATS["styles"] > 0:
        backup_src()                    # 确有改动才备份(保存前)
        wb.save(SRC)
        print(f"[保存] {SRC} (改动: 单元格值 {STATS['cells']} 处, 样式 {STATS['styles']} 处)")
        wb2 = openpyxl.load_workbook(SRC, data_only=True)
        ws_v = wb2[SHEET]
    else:
        print("[保存] 数据与格式均无变化, 未改写文件(未生成备份)")
        ws_v = ws

    n_rows, sort_viol, mono_viol = verify(ws_v)
    if sort_viol or mono_viol:
        sys.exit(f"校验未通过: 排序违规 {sort_viol}, 单调性违规 {mono_viol}")
    print(f"同步完成 ✔ 总耗时 {time.time() - t0:.1f} 秒")


if __name__ == "__main__":
    main()
