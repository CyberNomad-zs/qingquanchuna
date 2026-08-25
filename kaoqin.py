# -*- coding: utf-8 -*-
"""
考勤报表 → 考勤整理结果  自动整理脚本

输入: {月}月考勤报表_{起}_{止}.xlsx   (钉钉/考勤系统导出的宽表)
输出: {月}月考勤整理结果.xlsx         (三张表，复刻现有 7月结果表样式)

规则 (与 WB—考勤规则.docx / 虾的考勤—考勤核对规则.docx 一致):
  假期总时长 = 事假(小时) + 调休假(小时)
  净事假/调休 = MAX(0, 假期 - 加班)
  净加班     = MAX(0, 加班 - 假期)        # 两者互斥
  实际出勤   = 应出勤 - 净事假/调休/7.5 - 年假 - 病假 - 婚假 - 育儿假  (保留小数, 不四舍五入)
  迟到次数   直接抄报表

输出三张表:
  SHEET1 考勤整理结果      : 在职 且 (有请假/加班 或 当月入职) 的人 (全自动)
  SHEET2 打卡记录异常说明  : 有异常指标的人 (指标自动, 定性列留空标黄待人工补)
  SHEET3 其他员工(无请假加班): 无请假加班 但有异常/离职 的人 (全自动; 当月入职在职已并入SHEET1)

用法:
  python kaoqin.py 7月考勤报表_20260701_20260731.xlsx
  python kaoqin.py 8月考勤报表_20260801_20260831.xlsx --month 8
  python kaoqin.py <输入.xlsx> --month 7 --year 2026 --check 7月考勤整理结果.xlsx
"""
import argparse
import os
import re
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------- 样式常量 (复刻 7月结果表) ----------------
FONT_NAME = "微软雅黑"
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="00333333")
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True, color="00FFFFFF")
DATA_FONT = Font(name=FONT_NAME, size=10, bold=False)
RULE_TITLE_FONT = Font(name=FONT_NAME, size=12, bold=True)
RULE_BODY_FONT = Font(name=FONT_NAME, size=10, bold=False)

HEADER_FILL = PatternFill("solid", fgColor="004472C4")   # 表头蓝底
YELLOW_FILL = PatternFill("solid", fgColor="00FFF2CC")   # SHEET3 数据黄底 / SHEET2 待补黄底
HIGHLIGHT_FILL = PatternFill("solid", fgColor="00FFEB9C")  # SHEET2 定性留空标黄(偏橙, 区别于数据黄底)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HOURS_PER_DAY = 7.5

# ---------------- 输入列名 → 列索引 (按表头名定位, 鲁棒于列顺序) ----------------
INPUT_COL_NAMES = {
    "姓名": None, "人员类型": None, "员工状态": None, "入职日期": None, "离职日期": None,
    "应出勤天数": None, "工作日出勤天数": None, "补卡次数": None, "迟到次数": None,
    "上班缺卡次数": None, "下班缺卡次数": None, "旷工天数": None,
    "婚假(天)": None, "事假(小时)": None, "育儿假(天)": None, "病假(天)": None,
    "调休假(小时)": None, "年假(天)": None, "加班总时长(小时)": None,
}

# ---------------- 各 sheet 列定义 ----------------
SHEET1_HEADERS = ["姓名", "应出勤天数", "事假(小时)", "调休假(小时)", "加班总时长(小时)",
                  "年假(天)", "病假(天)", "婚假(天)", "育儿假(天)",
                  "净事假/调休(小时)", "净加班(小时)", "实际出勤天数", "迟到次数", "备注"]
SHEET1_WIDTHS = [10, 12, 12, 12, 14, 10, 10, 10, 10, 14, 12, 14, 10, 50]

SHEET2_HEADERS = ["姓名", "异常类型", "异常详情", "补卡次数", "缺卡次数(上/下)", "备注"]
SHEET2_WIDTHS = [10, 14, 55, 10, 12, 25]

SHEET3_HEADERS = ["姓名", "人员类型", "员工状态", "应出勤天数", "工作日出勤天数", "迟到次数", "备注"]
SHEET3_WIDTHS = [10, 10, 10, 12, 14, 10, 40]

RULE_LINES = [
    "计算规则说明",
    "1. 假期总时长 = 事假(小时) + 调休假(小时)",
    "2. 假期总时长与加班总时长互抵：加班≥假期→净事假/调休=0,净加班=加班-假期；加班<假期→净加班=0,净事假/调休=假期-加班",
    "3. 7.5小时 = 1天",
    "4. 实际出勤天数 = 应出勤 - 净事假/调休÷7.5 - 年假 - 病假 - 婚假 - 育儿假（保留小数，不四舍五入）",
    "5. 迟到次数直接抄报表，不修改",
]


def fnum(v):
    """把单元格值转成 float, 空值/非数 → 0.0"""
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def f1(x):
    """统一数字格式: 3.5→'3.5', 6→'6.0' (与现有结果表备注格式一致)"""
    return f"{x:.1f}"


def detect_month_year(text, default_year=2026):
    """从文件名里抓 月/年, 如 '7月考勤报表_20260701_20260731' → (2026, 7)"""
    m = re.search(r"(\d{4})(\d{2})(\d{2})", text)
    if m:
        y = int(m.group(1)); mo = int(m.group(2))
        if 1 <= mo <= 12:
            return y, mo
    m = re.search(r"(\d{1,2})月", text)
    if m:
        mo = int(m.group(1))
        if 1 <= mo <= 12:
            ym = re.search(r"(\d{4})年", text)
            return (int(ym.group(1)) if ym else default_year), mo
    return default_year, None


def read_input(path):
    """读输入宽表, 返回 (list[dict], year, month). 每条 dict 含需要的字段。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws = wb.active
    # 第1行分组表头(合并), 第2行字段名。姓名是 A 列合并单元格(A1:A2),
    # 第2行里没有"姓名"字样, 所以固定取列1。
    header = {"姓名": 1}
    for c in range(2, ws.max_column + 1):
        name = ws.cell(row=2, column=c).value
        if name:
            header[str(name).strip()] = c
    # 校验必须的列都在
    missing = [k for k in INPUT_COL_NAMES if k not in header]
    if missing:
        sys.exit(f"[错误] 输入表缺少列: {missing}\n请确认导出的考勤报表是否包含这些字段。")
    idx = {k: header[k] for k in INPUT_COL_NAMES}

    employees = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=idx["姓名"]).value
        if not name or str(name).strip() == "":
            continue
        def g(key):
            return ws.cell(row=r, column=idx[key]).value
        employees.append({
            "姓名": str(name).strip(),
            "人员类型": str(g("人员类型") or "").strip(),
            "员工状态": str(g("员工状态") or "").strip(),
            "入职日期": g("入职日期"),
            "离职日期": g("离职日期"),
            "应出勤天数": fnum(g("应出勤天数")),
            "工作日出勤天数": fnum(g("工作日出勤天数")),
            "补卡次数": fnum(g("补卡次数")),
            "迟到次数": fnum(g("迟到次数")),
            "上班缺卡次数": fnum(g("上班缺卡次数")),
            "下班缺卡次数": fnum(g("下班缺卡次数")),
            "旷工天数": fnum(g("旷工天数")),
            "婚假(天)": fnum(g("婚假(天)")),
            "事假(小时)": fnum(g("事假(小时)")),
            "育儿假(天)": fnum(g("育儿假(天)")),
            "病假(天)": fnum(g("病假(天)")),
            "调休假(小时)": fnum(g("调休假(小时)")),
            "年假(天)": fnum(g("年假(天)")),
            "加班总时长(小时)": fnum(g("加班总时长(小时)")),
        })
    wb.close()

    year, month = detect_month_year(os.path.basename(path))
    return employees, year, month


# ---------------- 计算 ----------------
def compute(emp):
    """对单个员工计算净事假/净加班/实际出勤/备注。返回 dict。"""
    shijia = emp["事假(小时)"]
    tiaoxiu = emp["调休假(小时)"]
    jiaban = emp["加班总时长(小时)"]
    jiaqi = shijia + tiaoxiu
    net_leave = max(0.0, jiaqi - jiaban)   # 净事假/调休
    net_ot = max(0.0, jiaban - jiaqi)       # 净加班
    actual = (emp["应出勤天数"] - net_leave / HOURS_PER_DAY
              - emp["年假(天)"] - emp["病假(天)"]
              - emp["婚假(天)"] - emp["育儿假(天)"])
    # 备注格式: 事假X + 调休假Y + 加班Z → 净A、净B  (与现有结果表一致)
    parts = []
    if shijia > 0: parts.append(f"事假{f1(shijia)}小时")
    if tiaoxiu > 0: parts.append(f"调休假{f1(tiaoxiu)}小时")
    if jiaban > 0: parts.append(f"加班{f1(jiaban)}小时")
    src = " + ".join(parts) if parts else "无"
    remark = (f"{src} → 净事假/调休{f1(net_leave)}小时、净加班{f1(net_ot)}小时")
    return {
        "净事假/调休(小时)": net_leave,
        "净加班(小时)": net_ot,
        "实际出勤天数": actual,
        "备注": remark,
    }


def has_leave_or_ot(emp):
    return (emp["事假(小时)"] > 0 or emp["调休假(小时)"] > 0 or emp["加班总时长(小时)"] > 0)


def is_resigned(emp):
    """离职: 员工状态含'离职' 或 离职日期非空"""
    if "离职" in emp["员工状态"]:
        return True
    rd = emp["离职日期"]
    return rd not in (None, "", "-")


def has_anomaly(emp):
    """有打卡异常指标: 旷工>0 或 缺卡>0 或 补卡>0 或 迟到>0"""
    return (emp["旷工天数"] > 0 or emp["上班缺卡次数"] > 0 or emp["下班缺卡次数"] > 0
            or emp["补卡次数"] > 0 or emp["迟到次数"] > 0)


def _date_ym(v):
    """从日期值取 (年, 月), 取不到返回 None。"""
    if v in (None, "", "-"):
        return None
    if hasattr(v, "year") and hasattr(v, "month"):
        return (v.year, v.month)
    m = re.search(r"(\d{4})[-/.]?(\d{1,2})[-/.]?(\d{1,2})", str(v))
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return None


def is_join_this_month(emp, year, month):
    """入职日期落在当月"""
    ym = _date_ym(emp["入职日期"])
    return ym is not None and ym == (year, month)


def has_join_or_leave(emp, year=None, month=None):
    """入职当月 或 离职(任何时间, 离职即不核算)。
    入职判断需 year/month; 不传则只看离职。"""
    if is_resigned(emp):
        return True
    if year is not None and month is not None and is_join_this_month(emp, year, month):
        return True
    return False


def fmt_date(v):
    """日期统一成 YYYY-MM-DD; 已经是字符串就规整, datetime 取 date。"""
    if v in (None, "", "-"):
        return ""
    s = str(v)
    # 处理 datetime.date / datetime.datetime
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    m = re.search(r"(\d{4})[-/.]?(\d{1,2})[-/.]?(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return s


# ---------------- 写样式工具 ----------------
def style_title(cell):
    cell.font = TITLE_FONT; cell.alignment = CENTER

def style_header(cell):
    cell.font = HEADER_FONT; cell.fill = HEADER_FILL; cell.alignment = CENTER; cell.border = BORDER

def style_data(cell, align=CENTER, fill=None):
    cell.font = DATA_FONT; cell.alignment = align; cell.border = BORDER
    if fill: cell.fill = fill

def style_rule(cell, bold=False):
    cell.font = RULE_TITLE_FONT if bold else RULE_BODY_FONT
    cell.alignment = LEFT


def write_sheet1(ws, rows, year, month, excluded):
    """SHEET1 考勤整理结果: 标题 + 表头 + 数据 + 规则说明 + 排除名单"""
    # 列宽
    for i, w in enumerate(SHEET1_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # 标题 A1:L1
    ws.merge_cells("A1:L1")
    ws["A1"] = f"{year}年{month}月考勤整理结果"
    style_title(ws["A1"])
    # 表头 第3行
    for i, h in enumerate(SHEET1_HEADERS, 1):
        c = ws.cell(row=3, column=i, value=h); style_header(c)
    # 数据 第4行起
    r = 4
    for emp in rows:
        comp = compute(emp)
        # 当月入职在职的新人: 备注补一句入职日
        # (钉钉应出勤已按入职后算, 实际出勤用该值即可, 这里只标注身份)
        if is_join_this_month(emp, year, month) and not is_resigned(emp):
            jd = fmt_date(emp["入职日期"])
            if comp["备注"].startswith("无"):
                comp["备注"] = f"{jd}入职，按入职日起算正常出勤"
            else:
                comp["备注"] += f"｜{jd}入职，按入职日起算"
        vals = [emp["姓名"], emp["应出勤天数"], emp["事假(小时)"], emp["调休假(小时)"],
                emp["加班总时长(小时)"], emp["年假(天)"], emp["病假(天)"], emp["婚假(天)"],
                emp["育儿假(天)"], comp["净事假/调休(小时)"], comp["净加班(小时)"],
                comp["实际出勤天数"], emp["迟到次数"], comp["备注"]]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=i, value=v)
            style_data(cell, align=LEFT if i == 14 else CENTER)
        r += 1
    # 规则说明区 (空一行)
    rule_start = r + 1
    for k, line in enumerate(RULE_LINES):
        rr = rule_start + k
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=14)
        c = ws.cell(row=rr, column=1, value=line)
        style_rule(c, bold=(k == 0))
    # 排除名单行
    excl_row = rule_start + len(RULE_LINES)
    ws.merge_cells(start_row=excl_row, start_column=1, end_row=excl_row, end_column=14)
    c = ws.cell(row=excl_row, column=1, value=f"已排除人员：{excluded}")
    style_rule(c)
    ws.freeze_panes = "A4"


def write_sheet2(ws, rows, year, month):
    """SHEET2 打卡记录异常说明: 指标自动, 异常类型/详情/备注 留空标黄待人工补"""
    for i, w in enumerate(SHEET2_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.merge_cells("A1:F1")
    ws["A1"] = "打卡记录异常检查"
    style_title(ws["A1"])
    for i, h in enumerate(SHEET2_HEADERS, 1):
        style_header(ws.cell(row=3, column=i, value=h))
    # 每个异常员工可能有多条(原表人工分多条), 这里按指标自动生成可机械确定的行:
    #   迟到行(若迟到>0) / 缺卡行(若缺卡>0) / 旷工行(若旷工>0)
    # 异常类型填机械值, 异常详情/备注留空标黄
    r = 4
    for emp in rows:
        bk = emp["补卡次数"]; up = emp["上班缺卡次数"]; dn = emp["下班缺卡次数"]
        late = emp["迟到次数"]; kuang = emp["旷工天数"]
        blocks = []
        if late > 0:
            blocks.append(("迟到", f"迟到{int(late)}次", "", ""))
        if up > 0 or dn > 0:
            blocks.append(("缺卡", f"上班缺卡{int(up)}次, 下班缺卡{int(dn)}次", "", f"补卡{int(bk)}次"))
        if kuang > 0:
            blocks.append(("旷工", f"旷工{f1(kuang)}天", "", "严重异常" if kuang >= emp["应出勤天数"] else "需关注"))
        if not blocks:  # 入离职等无指标但有异常标记的, 留一行空待补
            blocks.append(("", "", "", ""))
        for btype, detail, _x, note in blocks:
            ws.cell(row=r, column=1, value=emp["姓名"])
            # 异常类型(机械可定): 迟到/缺卡/旷工 填实, 其它留空标黄
            c2 = ws.cell(row=r, column=2, value=btype)
            ws.cell(row=r, column=3, value=detail)   # 异常详情: 机械部分填(迟到X次/缺卡X次/旷工X天)
            ws.cell(row=r, column=4, value=int(bk) if (up > 0 or dn > 0) else (int(bk) if bk else 0))
            ws.cell(row=r, column=5, value=f"{int(up)}/{int(dn)}")
            ws.cell(row=r, column=6, value=note)
            for i in range(1, 7):
                style_data(ws.cell(row=r, column=i), align=LEFT if i in (3, 6) else CENTER)
            # 机械填不了的人工列(异常类型为空 → 全行标黄; 详情/备注里人工补充部分用高亮提示)
            if not btype:
                for i in range(1, 7):
                    ws.cell(row=r, column=i).fill = HIGHLIGHT_FILL
            r += 1
    ws.freeze_panes = "A4"


def write_sheet3(ws, rows, year, month):
    """SHEET3 其他员工(无请假加班): 全自动, 黄底"""
    for i, w in enumerate(SHEET3_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.merge_cells("A1:G1")
    ws["A1"] = "无请假/加班数据员工汇总"
    style_title(ws["A1"])
    for i, h in enumerate(SHEET3_HEADERS, 1):
        style_header(ws.cell(row=3, column=i, value=h))
    r = 4
    for emp in rows:
        # 备注: 旷工X天, 缺卡Y次 [+ 入离职说明]
        parts = []
        kuang = emp["旷工天数"]
        if kuang > 0: parts.append(f"旷工{f1(kuang)}天")
        miss = int(emp["上班缺卡次数"] + emp["下班缺卡次数"])
        if miss > 0: parts.append(f"缺卡{miss}次")
        jd = fmt_date(emp["入职日期"]); rd = fmt_date(emp["离职日期"])
        full_absent = kuang >= emp["应出勤天数"] and emp["应出勤天数"] > 0
        if rd:
            parts.append(f"离职日期{rd}")
        elif is_join_this_month(emp, year, month):
            tag = "暂无加班" if emp["加班总时长(小时)"] == 0 else f"加班{f1(emp['加班总时长(小时)'])}小时"
            if full_absent:
                parts.append(f"{jd}入职，全月缺卡/旷工，可能未纳入考勤系统")
            else:
                parts.append(f"{jd}入职，按入职日起算正常出勤，{tag}")
        elif full_absent:
            ptype = emp["人员类型"]
            hint = "外包/实习人员，可能未纳入考勤系统" if ptype in ("外包", "实习") else "全月缺卡/旷工，需核查"
            parts.append(hint)
        remark = ", ".join(parts)
        vals = [emp["姓名"], emp["人员类型"], emp["员工状态"], emp["应出勤天数"],
                emp["工作日出勤天数"], emp["迟到次数"], remark]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c, align=LEFT if i == 7 else CENTER, fill=YELLOW_FILL)
        r += 1
    ws.freeze_panes = "A4"


def build_excluded_list(all_emps, sheet1_names, sheet3_names):
    """排除名单 = 全员 - SHEET1 - SHEET3"""
    included = set(sheet1_names) | set(sheet3_names)
    excl = [e["姓名"] for e in all_emps if e["姓名"] not in included]
    # 离职的单独标注
    result = []
    for e in all_emps:
        if e["姓名"] in included:
            continue
        tag = "（离职）" if is_resigned(e) else ""
        result.append(f"{e['姓名']}{tag}")
    return "、".join(result)


# ---------------- 比对核验 ----------------
def compare_with_existing(generated_path, existing_path):
    """把生成的 SHEET1 和现有结果表逐列比对, 打印差异。"""
    if not os.path.exists(existing_path):
        print(f"[核验] 未找到现有结果表 {existing_path}, 跳过比对。")
        return
    g = openpyxl.load_workbook(generated_path, data_only=True)
    e = openpyxl.load_workbook(existing_path, data_only=True)
    gw = g["考勤整理结果"]; ew = e["考勤整理结果"]
    # 现有表数据行: 第4行起到规则说明前
    def read_rows(ws, header_row=3):
        headers = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
        rows = {}
        r = header_row + 1
        while True:
            name = ws.cell(row=r, column=1).value
            if name is None or str(name).strip() == "":
                break
            rows[str(name).strip()] = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
            r += 1
        return headers, rows
    gh, gr = read_rows(gw)
    eh, er = read_rows(ew)
    print(f"[核验] 生成 {len(gr)} 人, 现有 {len(er)} 人")
    only_g = set(gr) - set(er); only_e = set(er) - set(gr)
    if only_g: print(f"  [仅生成] {only_g}")
    if only_e: print(f"  [仅现有] {only_e}")
    # 逐列比对公共人员
    diff_count = 0
    for name in sorted(set(gr) & set(er)):
        for i, (a, b) in enumerate(zip(gr[name], er[name])):
            if i in (12, 13):  # 备注列(12=迟到? 重新核对: 列13=迟到,14=备注 → index12=迟到,13=备注)
                continue
            # 数值容差
            try:
                if abs(float(a) - float(b)) > 0.001:
                    print(f"  [差异] {name} 列{i+1}({gh[i]}): 生成={a} 现有={b}")
                    diff_count += 1
            except (TypeError, ValueError):
                if str(a) != str(b):
                    print(f"  [差异] {name} 列{i+1}({gh[i]}): 生成={a!r} 现有={b!r}")
                    diff_count += 1
    if diff_count == 0 and not only_g and not only_e:
        print("[核验] ✅ SHEET1 与现有结果表逐列一致(不含备注列)")
    else:
        print(f"[核验] ⚠️ 共 {diff_count} 处数值差异, 详见上方")
    g.close(); e.close()


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser(description="考勤报表 → 考勤整理结果")
    ap.add_argument("input", help="输入考勤报表 xlsx 路径")
    ap.add_argument("--month", type=int, help="月份(1-12), 不填则从文件名识别")
    ap.add_argument("--year", type=int, default=2026, help="年份, 默认 2026")
    ap.add_argument("--check", help="现有结果表路径, 用于逐列比对核验")
    ap.add_argument("-o", "--output", help="输出文件名, 不填则用 {月}月考勤整理结果_校验.xlsx")
    args = ap.parse_args()

    if not os.path.exists(args.input):
        sys.exit(f"[错误] 输入文件不存在: {args.input}")

    employees, det_year, det_month = read_input(args.input)
    year = args.year
    month = args.month if args.month else det_month
    if month is None:
        sys.exit("[错误] 无法从文件名识别月份, 请用 --month 指定 (如 --month 7)")
    print(f"[读取] {len(employees)} 人, 年月={year}年{month}月")

    # 分流规则 (与用户确认口径一致):
    #   SHEET1: 在职 且 (有请假/加班 或 当月入职)   ← 当月入职新人并入主表, 按入职后算出勤
    #   SHEET3: 无请假/加班 且 (有异常 或 离职)      ← 当月入职在职已并入SHEET1, 不再在此
    #   SHEET2: 所有有异常指标的人(供打卡异常留档)
    #   排除名单: 其余 = 全勤无请假加班的在职老员工 (+ 离职但无任何异常的人, 罕见)
    sheet1, sheet2, sheet3 = [], [], []
    for emp in employees:
        resigned = is_resigned(emp)
        join_this_month = is_join_this_month(emp, year, month)
        if not resigned and (has_leave_or_ot(emp) or join_this_month):
            # 在职 且 (有请假加班 或 当月入职) → 主表 S1
            sheet1.append(emp)
        elif (not has_leave_or_ot(emp)) and (has_anomaly(emp) or has_join_or_leave(emp, year, month)):
            # 无请假加班, 且有异常/离职 → SHEET3 (当月入职在职已被上面接走)
            sheet3.append(emp)
        # SHEET2: 有异常指标的都记录(在职/离职都留档)
        if has_anomaly(emp):
            sheet2.append(emp)
    print(f"[分流] SHEET1={len(sheet1)} SHEET2={len(sheet2)} SHEET3={len(sheet3)}")

    excluded = build_excluded_list(employees, [e["姓名"] for e in sheet1], [e["姓名"] for e in sheet3])

    # 写出
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = "考勤整理结果"
    write_sheet1(ws1, sheet1, year, month, excluded)
    ws2 = wb.create_sheet("打卡记录异常说明")
    write_sheet2(ws2, sheet2, year, month)
    ws3 = wb.create_sheet("其他员工(无请假加班)")
    write_sheet3(ws3, sheet3, year, month)

    out = args.output or f"{month}月考勤整理结果_校验.xlsx"
    wb.save(out); wb.close()
    print(f"[输出] {out}")

    if args.check:
        compare_with_existing(out, args.check)


if __name__ == "__main__":
    main()
