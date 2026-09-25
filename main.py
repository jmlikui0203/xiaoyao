# -*- coding: utf-8 -*-
"""
超料记录 资料同步脚本（对应 AGENT.md 硬流程）

执行顺序（AGENT.md 头部约定）:
  §4  PG视图 production_start_finish 同步 指令/数量/三完成日
  §10 orders 表补全 型号/型体/颜色/交期 (work_no 组合号拆分 + norm 归一化)
  §11 特殊订单(无主指令数量) 分单数量 + 生产累计完成日
  §11.1 特殊订单 型号/型体/颜色/交期 (交期取主指令与分单全局最晚)
  §9  补料损耗表 按截止日汇总 超料1/2/3 (多文件取文件名日期最新)
  §5  全表格式统一为第 3 行基准
  §12 按 交期(列6) 从小到大排序(收尾)

运行: E:/tools/uv/uv.exe run python main.py  (在 E:/xiaoyao 下)
"""
import re
import shutil
import glob
import os
import sys
from datetime import datetime, date
from collections import defaultdict
from copy import copy

import openpyxl
import psycopg2

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

MARK_NODATA = "-"  # 超料列: 源表查无此指令(无数据), 区别于 0

VIEW_FILTER_FROM = "15D09A001"  # production/视图同口径过滤

norm = lambda s: "".join(str(s).split()).upper()          # 去所有空白+大写
SUB_PAT = re.compile(r"^(.+)-\d{2,}$")                    # 分单: xxx-NN
LOSS_FILE_PAT = re.compile(r"(\d{8})")                    # 文件名中的 8 位日期


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


def backup_src():
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    bak = rf"E:/xiaoyao/超料记录_备份_{ts}.xlsx"
    shutil.copy(SRC, bak)
    print(f"[备份] {bak}")
    return bak


def pick_loss_file():
    """§9: 多份补料损耗表时, 取文件名末尾 8 位日期最大的那一份(以文件名为准, 不看修改时间)"""
    files = []
    for pat in (r"E:/xiaoyao/補料發料損耗表-*.xlsx", r"E:/xiaoyao/补料发料损耗表-*.xlsx"):
        files += glob.glob(pat)
    if not files:
        raise FileNotFoundError("目录下未找到 補料發料損耗表-*.xlsx / 补料发料损耗表-*.xlsx")
    dated = []
    for f in files:
        m = LOSS_FILE_PAT.search(os.path.basename(f))
        if m:
            dated.append((m.group(1), f))
    if not dated:
        raise ValueError(f"补料损耗表文件名中提取不到 8 位日期: {files}")
    dated.sort(key=lambda x: x[0], reverse=True)
    if len(dated) > 1 and dated[0][0] == dated[1][0]:
        raise ValueError(f"同一天存在多份补料损耗表, 需人工确认: {dated}")
    print(f"[§9] 补料损耗表共 {len(files)} 份, 取最新: {dated[0][1]} (日期 {dated[0][0]})")
    return dated[0][1]


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
    """§4: 视图同步(命中更新/未命中追加); NULL 不写入不清空"""
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
            ws.cell(r, COL_INSTR).value = wno
            appended += 1
        for col, val in vals:
            if val is not None:  # NULL 不写入不清空
                c = ws.cell(r, col)
                c.value = to_datetime(val) if col != COL_QTY else val
                if col != COL_QTY:
                    c.number_format = DATE_FMT
    print(f"[§4] 视图同步: 更新 {updated} 行, 追加 {appended} 行")


def step10_orders_fill(ws, comp_map):
    """§10: 型号/型体/颜色/交期 (组合号已拆分, norm 匹配; 交期冲突默认取最晚)"""
    filled = unmatched = 0
    cfm_conflicts = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, COL_INSTR).value
        if v is None:
            continue
        recs = comp_map.get(norm(v))
        if not recs:
            unmatched += 1
            continue
        d = recs[0]
        ws.cell(r, COL_MODEL).value = d["sno"]
        ws.cell(r, COL_STYLE).value = d["sname"]
        ws.cell(r, COL_COLOR).value = d["col"]
        cfms = [x["cfm"] for x in recs if x["cfm"] is not None]
        if len({str(x) for x in cfms}) > 1:
            cfm_conflicts.append((str(v), max(cfms)))  # 冲突默认取最晚(AGENT.md §10)
        if cfms:
            c = ws.cell(r, COL_CFM)
            c.value = to_datetime(max(cfms))
            c.number_format = DATE_FMT
        filled += 1
    print(f"[§10] orders 四列: 填充 {filled} 行, 未匹配 {unmatched} 行"
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
        if bare_qty.get(m, 0) > 0:      # 主指令有有效数量 -> 非特殊订单
            continue                     # (判定不能用"裸键不存在", 裸键可能 quantity=NULL)
        subs = {k: q for k, q in sub_qty.items()
                if k.startswith(m + "-") and SUB_PAT.match(k)}
        if not subs:
            continue
        n += 1
        # 数量 = 分单求和
        total = sum(subs.values())
        if total > 0:
            ws.cell(r, COL_QTY).value = total
        # 完成日 = 主指令生产累计 >= 分单合计 的最早日期(工序独立)
        cf, sf, af = finish_dates_per_proc(prod.get(m, []), total)
        for col, fd in ((COL_CUT, cf), (COL_STITCH, sf), (COL_ASM, af)):
            if fd is not None:
                c = ws.cell(r, col)
                c.value = to_datetime(fd)
                c.number_format = DATE_FMT
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
            ws.cell(r, COL_MODEL).value = sno
            ws.cell(r, COL_STYLE).value = sname
            ws.cell(r, COL_COLOR).value = col_
        # §11.1 交期 = max(主指令, 全部分单)
        all_cfm = [x["cfm"] for k in subs for x in comp_map[k] if x["cfm"] is not None]
        if direct and direct[0]["cfm"] is not None:
            all_cfm.append(direct[0]["cfm"])
        if all_cfm:
            c = ws.cell(r, COL_CFM)
            c.value = to_datetime(max(all_cfm))
            c.number_format = DATE_FMT
    print(f"[§11] 特殊订单(分单汇总): 处理 {n} 行")


def step9_loss_summary(ws, loss_file):
    """§9: 补料责任金额按三截止日累加 -> 超料1/2/3 (含 -/0/空 三态语义)"""
    lw = openpyxl.load_workbook(loss_file, data_only=True)
    lws = lw[lw.sheetnames[0]]
    loss = defaultdict(list)
    for r in range(2, lws.max_row + 1):
        ino = lws.cell(r, 7).value
        if ino is None:
            continue
        amt = lws.cell(r, 17).value
        if amt is None:
            continue
        try:
            amt = float(amt)
        except (TypeError, ValueError):
            continue
        d = pdate(lws.cell(r, 1).value)   # 日期 NULL/无法解析 -> 跳过该行
        if d is None:
            continue
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
                ws.cell(r, col_a).value = None
                n_blank += 1
            elif rows:                               # 源表内 -> 累计和(含0)
                ws.cell(r, col_a).value = round(
                    sum(a for d, a in rows if d <= cutoff), 2)
                n_num += 1
            else:                                    # 源表查无 -> 标记
                ws.cell(r, col_a).value = MARK_NODATA
                n_mark += 1
    print(f"[§9] 超料1/2/3 汇总: 数值 {n_num} | 标记'-' {n_mark} | 完成日空留空 {n_blank}")


def step5_unify_style(ws):
    """§5: 全部数据行统一为第 3 行原始样式(用户否决第 2 行)"""
    maxc = ws.max_column
    styles = {}
    for c in range(1, maxc + 1):
        cell = ws.cell(3, c)
        styles[c] = (cell.font, cell.fill, cell.border,
                     cell.alignment, cell.number_format, cell.protection)
    height = ws.row_dimensions[3].height
    for r in range(2, ws.max_row + 1):
        if height is not None:
            ws.row_dimensions[r].height = height
        for c in range(1, maxc + 1):
            cell = ws.cell(r, c)
            f, fi, b, a, nf, p = styles[c]
            cell.font = copy(f); cell.fill = copy(fi); cell.border = copy(b)
            cell.alignment = copy(a); cell.number_format = nf; cell.protection = copy(p)
    print(f"[§5] 格式统一: 第 3 行基准 -> {ws.max_row - 1} 行")


def step12_sort_by_cfm(ws):
    """§12: 按交期升序(整行移动), 空值排末尾(稳定排序)"""
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

    rows_sorted = sorted(rows, key=key)   # list.sort 稳定
    for i, vals in enumerate(rows_sorted, start=2):
        for c, v in enumerate(vals, start=1):
            ws.cell(i, c).value = v
    n_with = sum(1 for vals in rows_sorted if key(vals)[0] == 0)
    print(f"[§12] 交期排序: {len(rows_sorted)} 行 (有交期 {n_with}, 空排末尾 {len(rows_sorted) - n_with})")


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
    if not os.path.exists(SRC):
        sys.exit(f"目标文件不存在: {SRC}")
    backup_src()
    loss_file = pick_loss_file()
    view, comp_map, sub_qty, bare_qty, prod = fetch_pg_data()

    wb = openpyxl.load_workbook(SRC)
    ws = wb[SHEET]

    step4_sync_view(ws, view)
    step10_orders_fill(ws, comp_map)
    step11_special_orders(ws, comp_map, sub_qty, bare_qty, prod)
    step9_loss_summary(ws, loss_file)
    step5_unify_style(ws)
    step12_sort_by_cfm(ws)

    wb.save(SRC)
    print(f"[保存] {SRC}")

    wb2 = openpyxl.load_workbook(SRC, data_only=True)
    n_rows, sort_viol, mono_viol = verify(wb2[SHEET])
    if sort_viol or mono_viol:
        sys.exit(f"校验未通过: 排序违规 {sort_viol}, 单调性违规 {mono_viol}")
    print("同步完成 ✔")


if __name__ == "__main__":
    main()
