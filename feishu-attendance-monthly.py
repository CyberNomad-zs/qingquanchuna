"""每月考勤数据上传飞书 · 可重复、换月不改代码。

把当月的:
  - 钉钉考勤原始报表(*月考勤报表_起_止.xlsx) → 「{月}考勤报表」(57列) + 「{月}打卡时间明细」(10列)
  - kaoqin.py 整理结果(*月考勤整理结果_校验.xlsx)的 3 个 sheet →
      Sheet1 考勤整理结果   → 「{月}考勤结果表」(14列)
      Sheet2 打卡记录异常说明 → 「{月}打卡异常说明」(6列)
      Sheet3 其他员工        → 「{月}其他员工」(7列)

每月在 base Xp7UbFxhRavG7OsxAP3cnAfXnJh 里新建 5 张「{月}」前缀表(已存在则复用),
清空该表后灌入当月数据。历史月表保留不动。

零依赖:仅用 stdlib(urllib + openpyxl)。凭据直接读 .env,不耦合 backend。
打卡时间用北京时间 +08:00,飞书日期字段存绝对毫秒按用户时区回显。

用法:
  python feishu-attendance-monthly.py <报表.xlsx> <整理结果.xlsx> [--apply]
  python feishu-attendance-monthly.py 7月考勤报表_xxx.xlsx 7月考勤整理结果_校验.xlsx   # 干跑
  python feishu-attendance-monthly.py 7月考勤报表_xxx.xlsx 7月考勤整理结果_校验.xlsx --apply
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone, timedelta
from pathlib import Path

import openpyxl

# ---------------- 常量 ----------------
BASE = "https://open.feishu.cn/open-apis"
APP_TOKEN = "Xp7UbFxhRavG7OsxAP3cnAfXnJh"
ENV_PATH = Path(__file__).resolve().parent / ".env"
CN_TZ = timezone(timedelta(hours=8))
BATCH = 500

# 飞书字段 type: 1=文本 2=数字 4=单选 5=日期 7=复选框 11=人员
# (注意:复选框真实 type=7,3 是旧/兼容编码,以总表实测 7 为准)
T_TEXT, T_NUM, T_CHECK, T_DATE, T_USER = 1, 2, 7, 5, 11

# ---------------- 5 张表字段定义(建表用) ----------------
# 57 列考勤报表:11 文本 + 1 日期(入职) + 45 数字(离职日期是文本型!)
SUMMARY_FIELDS = [
    ("姓名", T_TEXT), ("工号", T_TEXT), ("部门", T_TEXT), ("组织架构", T_TEXT),
    ("性别", T_TEXT), ("职务", T_TEXT), ("人员类型", T_TEXT), ("员工状态", T_TEXT),
    ("合同公司", T_TEXT), ("工作地点", T_TEXT), ("入职日期", T_DATE), ("离职日期", T_TEXT),
    ("考勤组", T_TEXT),
    ("应出勤天数", T_NUM), ("应出勤时长(小时)", T_NUM), ("工作日出勤天数", T_NUM),
    ("休息日出勤天数", T_NUM), ("休息或未排班天数", T_NUM), ("实际出勤时长(小时)", T_NUM),
    ("班内工作时长(小时)", T_NUM), ("出差时长", T_NUM), ("外出时长", T_NUM),
    ("外勤次数", T_NUM), ("补卡次数", T_NUM), ("迟到次数", T_NUM), ("迟到时长(小时)", T_NUM),
    ("严重迟到次数", T_NUM), ("严重迟到时长(小时)", T_NUM), ("早退次数", T_NUM),
    ("早退时长(小时)", T_NUM), ("缺勤时长(小时)", T_NUM), ("上班缺卡次数", T_NUM),
    ("下班缺卡次数", T_NUM), ("旷工天数", T_NUM),
    ("婚假(天)", T_NUM), ("产假(天)", T_NUM), ("哺乳假(小时)", T_NUM),
    ("事假(小时)", T_NUM), ("陪产假(天)", T_NUM), ("外出假(小时)", T_NUM),
    ("育儿假(天)", T_NUM), ("病假(天)", T_NUM), ("调休假(小时)", T_NUM),
    ("年假(天)", T_NUM), ("丧假(天)", T_NUM),
    ("加班总时长(小时)", T_NUM), ("加班总时长 - 计加班费(小时)", T_NUM),
    ("加班总时长 - 计调休(小时)", T_NUM), ("工作日加班(小时)", T_NUM),
    ("工作日加班 - 计加班费(小时)", T_NUM), ("工作日加班 - 计调休(小时)", T_NUM),
    ("休息日加班(小时)", T_NUM), ("休息日加班 - 计加班费(小时)", T_NUM),
    ("休息日加班 - 计调休(小时)", T_NUM), ("节假日加班(小时)", T_NUM),
    ("节假日加班 - 计加班费(小时)", T_NUM), ("节假日加班 - 计调休(小时)", T_NUM),
]
# 10 列打卡明细:2 文本 + 3 日期 + 2 复选 + 3 文本
DETAIL_FIELDS = [
    ("姓名", T_TEXT), ("工号", T_TEXT), ("日期", T_DATE),
    ("上班打卡时间", T_DATE), ("下班打卡时间", T_DATE),
    ("是否缺卡（早）", T_CHECK), ("是否缺卡（晚）", T_CHECK),
    ("工时", T_TEXT), ("迟到情况", T_TEXT), ("加班情况", T_TEXT),
]
# 17 列考勤结果表(对齐 kaoqin.py Sheet1 + 飞书手加的 3 列):
#   首列 姓名(文本,索引) + 姓名(人员)(type=11,飞书可@人,值来自总表映射)
#   + 月份(文本,当月值如 "2026-7") + 12 数字 + 备注(文本) + 是否确认(复选框 type=7,留空手勾)
# 注:飞书人员类型不能当首列,故姓名仍文本当索引,人员列紧跟其后。
RESULT_FIELDS = [
    ("姓名", T_TEXT),
    ("姓名 (人员 )", T_USER),
    ("月份", T_TEXT),
    ("应出勤天数", T_NUM), ("事假(小时)", T_NUM),
    ("调休假(小时)", T_NUM), ("加班总时长(小时)", T_NUM), ("年假(天)", T_NUM),
    ("病假(天)", T_NUM), ("婚假(天)", T_NUM), ("育儿假(天)", T_NUM),
    ("净事假/调休(小时)", T_NUM), ("净加班(小时)", T_NUM),
    ("实际出勤天数", T_NUM), ("迟到次数", T_NUM), ("备注", T_TEXT),
    ("是否确认", T_CHECK),
]
# 6 列打卡异常说明(对齐 Sheet2):全文本(缺卡次数(上/下)是 "5/5" 格式)
ANOMALY_FIELDS = [
    ("姓名", T_TEXT), ("异常类型", T_TEXT), ("异常详情", T_TEXT),
    ("补卡次数", T_TEXT), ("缺卡次数(上/下)", T_TEXT), ("备注", T_TEXT),
]
# 7 列其他员工(对齐 Sheet3):4 文本 + 3 数字
OTHER_FIELDS = [
    ("姓名", T_TEXT), ("人员类型", T_TEXT), ("员工状态", T_TEXT),
    ("应出勤天数", T_NUM), ("工作日出勤天数", T_NUM), ("迟到次数", T_NUM),
    ("备注", T_TEXT),
]

# spec_key → (字段定义, 源数据来源)
TABLE_SPECS = {
    "考勤报表":      SUMMARY_FIELDS,
    "打卡时间明细":  DETAIL_FIELDS,
    "考勤结果表":    RESULT_FIELDS,
    "打卡异常说明":  ANOMALY_FIELDS,
    "其他员工":      OTHER_FIELDS,
}

# 姓名→open_id 映射的权威源:用户手动维护「姓名(人员)」列的总表。
# 每月新月表的人员列值从这里读。改名(如 X月考勤结果表)不影响,按此常量名匹配。
PERSON_MAP_TABLE = "考勤结果总表"
PERSON_FIELD = "姓名 (人员 )"
PERSON_NAME_FIELD = "姓名"
# 总表累积:写入时跳过人工字段(「是否确认」由用户在飞书手勾,脚本绝不覆盖)
MASTER_IGNORE = {"是否确认"}

# 文本列名集合(写入时按文本处理;其余按数值)
SUMMARY_TEXT_COLS = {n for n, t in SUMMARY_FIELDS if t == T_TEXT}
SUMMARY_DATE_COLS = {n for n, t in SUMMARY_FIELDS if t == T_DATE}  # 入职日期
# 离职日期虽在文本列,但 _date_text 单独处理,下面 SUMMARY_TEXT_COLS 已含它


# ---------------- .env 与 token ----------------
def load_env() -> tuple[str, str]:
    app_id = app_secret = ""
    for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("FEISHU_APP_ID="):
            app_id = line.split("=", 1)[1].strip()
        elif line.startswith("FEISHU_APP_SECRET="):
            app_secret = line.split("=", 1)[1].strip()
    if not app_id or not app_secret:
        raise RuntimeError(f"未在 {ENV_PATH} 找到 FEISHU_APP_ID/SECRET")
    return app_id, app_secret


def api(method: str, path: str, body=None, token: str = ""):
    """飞书 API 调用,带网络层重试。

    网络层错误(URLError: SSL EOF / 连接拒绝 / 超时 / DNS)→ 指数退避重试 4 次
    (1s/2s/4s/8s)。飞书业务错误(HTTPError 带 JSON body,如 code≠0 的参数错、
    token 失效)不重试,直接返回 body 让上层按 code 判断。
    """
    url = f"{BASE}{path}"
    headers = {"Content-Type": "application/json; charset=utf-8"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    data = json.dumps(body).encode("utf-8") if body is not None else None

    max_retries = 4
    for attempt in range(max_retries + 1):
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            # 飞书业务错误:返回 body 让上层按 code 处理,不重试
            raw = b""
            try:
                raw = e.read()
            except Exception:  # noqa: BLE001
                pass
            try:
                return json.loads(raw.decode("utf-8"))
            except (ValueError, UnicodeDecodeError):
                # 无 JSON body(如裸 404):包装成统一错误结构
                return {"code": e.code, "msg": f"HTTP {e.code} {e.reason}",
                        "_no_body": True, "_url": url}
        except urllib.error.URLError as e:
            # 网络层错误:重试
            if attempt >= max_retries:
                raise
            wait = 2 ** attempt  # 1,2,4,8
            print(f"[网络] {method} {path} 失败({e}),{wait}s 后重试 "
                  f"({attempt + 1}/{max_retries})", flush=True)
            time.sleep(wait)
        except (TimeoutError, ConnectionResetError, OSError) as e:
            # 兜底:其他网络/系统层错误同样重试
            if attempt >= max_retries:
                raise
            wait = 2 ** attempt
            print(f"[网络] {method} {path} 失败({e}),{wait}s 后重试 "
                  f"({attempt + 1}/{max_retries})", flush=True)
            time.sleep(wait)


def get_token() -> str:
    aid, sec = load_env()
    d = api("POST", "/auth/v3/tenant_access_token/internal",
            {"app_id": aid, "app_secret": sec})
    if d.get("code") != 0:
        raise RuntimeError(f"token failed: {d}")
    return d["tenant_access_token"]


# ---------------- 飞书表管理 ----------------
def list_tables(token: str) -> dict[str, str]:
    """返回 {表名: table_id}。"""
    out: dict[str, str] = {}
    page = ""
    while True:
        path = f"/bitable/v1/apps/{APP_TOKEN}/tables?page_size=100"
        if page:
            path += f"&page_token={page}"
        d = api("GET", path, token=token)
        if d.get("code") != 0:
            raise RuntimeError(f"list_tables: {d}")
        for it in d.get("data", {}).get("items", []):
            out[it["name"]] = it["table_id"]
        if not d.get("data", {}).get("has_more"):
            break
        page = d["data"].get("page_token", "")
    return out


def load_person_map(token: str) -> dict[str, str]:
    """从「考勤结果总表」读 姓名→open_id 映射(用户在飞书手填的人员列)。

    每月新建「X月考勤结果表」时,人员列值查这份字典写入。总表不存在或无人员列
    则返回空字典(当月人员列全留空,日志提示去飞书手补)。
    """
    tables = list_tables(token)
    tid = tables.get(PERSON_MAP_TABLE)
    if not tid:
        print(f"[映射] 未找到总表「{PERSON_MAP_TABLE}」,人员列将全留空")
        return {}
    mapping: dict[str, str] = {}
    page = ""
    while True:
        path = (f"/bitable/v1/apps/{APP_TOKEN}/tables/{tid}/records"
                f"?page_size=500")
        if page:
            path += f"&page_token={page}"
        d = api("GET", path, token=token)
        if d.get("code") != 0:
            print(f"[映射] 读总表失败 code={d.get('code')},人员列将全留空")
            return {}
        data = d.get("data") or {}
        for it in data.get("items") or []:
            fields = it.get("fields", {}) or {}
            name = fields.get(PERSON_NAME_FIELD)
            person = fields.get(PERSON_FIELD)
            if name and isinstance(person, list) and person:
                pid = person[0].get("id")
                if pid:
                    mapping[str(name).strip()] = pid
        if not data.get("has_more"):
            break
        page = data.get("page_token", "")
    print(f"[映射] 从「{PERSON_MAP_TABLE}」读到 {len(mapping)} 人 姓名→open_id")
    return mapping


def create_table(token: str, name: str, fields: list[tuple[str, int]]) -> str:
    """建表(一次带 fields)。返回 table_id。"""
    body = {
        "table": {
            "name": name,
            "default_view_name": "默认视图",
            "fields": [{"field_name": n, "type": t} for n, t in fields],
        }
    }
    d = api("POST", f"/bitable/v1/apps/{APP_TOKEN}/tables", body, token)
    if d.get("code") != 0:
        raise RuntimeError(f"建表「{name}」失败: {d}")
    return d["data"]["table_id"]


def ensure_table(token: str, month_prefix: str, spec_key: str,
                 fields: list[tuple[str, int]], existing: dict[str, str]) -> str:
    """按「{月}{spec}」名找表;没有就建。返回 table_id。"""
    tname = f"{month_prefix}{spec_key}"
    if tname in existing:
        print(f"  表「{tname}」已存在,复用 {existing[tname]}")
        return existing[tname]
    tid = create_table(token, tname, fields)
    print(f"  新建表「{tname}」{len(fields)}字段 → {tid}")
    return tid


def list_record_ids(token: str, table_id: str) -> list[str]:
    ids: list[str] = []
    page = ""
    while True:
        path = f"/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/records?page_size=500"
        if page:
            path += f"&page_token={page}"
        d = api("GET", path, token=token)
        if d.get("code") != 0:
            raise RuntimeError(f"list records {table_id}: {d}")
        data = d.get("data") or {}
        for it in data.get("items") or []:
            ids.append(it["record_id"])
        if not data.get("has_more"):
            break
        page = data.get("page_token", "")
    return ids


def batch_delete(token: str, table_id: str, ids: list[str]) -> None:
    for i in range(0, len(ids), BATCH):
        chunk = ids[i:i + BATCH]
        d = api("POST", f"/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/records/batch_delete",
                {"records": chunk}, token)
        if d.get("code") != 0:
            raise RuntimeError(f"batch_delete {table_id}: {d}")
        print(f"    删除 {len(chunk)} 条 (累计 {i + len(chunk)}/{len(ids)})")


def batch_create(token: str, table_id: str, records: list[dict]) -> int:
    total = 0
    for i in range(0, len(records), BATCH):
        chunk = records[i:i + BATCH]
        d = api("POST",
                f"/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/records/batch_create",
                {"records": [{"fields": f} for f in chunk]}, token)
        if d.get("code") != 0:
            raise RuntimeError(f"batch_create {table_id} 批{i//BATCH+1}: {d}")
        total += len(chunk)
        print(f"    写入 {len(chunk)} 条 (累计 {total}/{len(records)})")
    return total


def batch_update(token: str, table_id: str, records: list[dict]) -> int:
    """records: [{"record_id":..., "fields": {...}}, ...]。"""
    total = 0
    for i in range(0, len(records), BATCH):
        chunk = records[i:i + BATCH]
        d = api("POST",
                f"/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/records/batch_update",
                {"records": chunk}, token)
        if d.get("code") != 0:
            raise RuntimeError(f"batch_update {table_id} 批{i//BATCH+1}: {d}")
        total += len(chunk)
        print(f"    更新 {len(chunk)} 条 (累计 {total}/{len(records)})")
    return total


def read_master_index(token: str, table_id: str) -> dict[tuple[str, str], str]:
    """读「考勤结果总表」现有 (姓名, 月份) → record_id,用于去重 upsert。

    月份文本列(type=1)值如 "2026-7"。
    """
    index: dict[tuple[str, str], str] = {}
    page = ""
    while True:
        path = (f"/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/records"
                f"?page_size=500")
        if page:
            path += f"&page_token={page}"
        d = api("GET", path, token=token)
        if d.get("code") != 0:
            raise RuntimeError(f"read_master_index {table_id}: {d}")
        data = d.get("data") or {}
        for it in data.get("items") or []:
            fields = it.get("fields", {}) or {}
            name = fields.get(PERSON_NAME_FIELD)
            mon = fields.get("月份")
            if name and mon is not None:
                index[(str(name).strip(), str(mon).strip())] = it["record_id"]
        if not data.get("has_more"):
            break
        page = data.get("page_token", "")
    return index


def sync_master(token: str, table_id: str, records: list[dict], month_value: str) -> int:
    """覆盖式写入「考勤结果总表」——先清空全部旧数据,再全量写入当月结果。"""
    if not records:
        return 0
    index = read_master_index(token, table_id)
    old_ids = list(index.values())
    if old_ids:
        print(f"  清空旧数据 {len(old_ids)} 条 …")
        batch_delete(token, table_id, old_ids)
    to_create = []
    for rec in records:
        name = rec.get(PERSON_NAME_FIELD)
        if not name:
            continue
        payload = {k: v for k, v in rec.items() if k not in MASTER_IGNORE}
        to_create.append(payload)
    created = batch_create(token, table_id, to_create) if to_create else 0
    return created


# ---------------- 值转换 ----------------
def _num(v):
    if v is None or v == "" or v == "-":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s == "-" else s


def _date_ms_utc(v):
    """入职日期/日期列(飞书日期型)→ UTC 00:00 毫秒。"""
    if v is None or v == "" or v == "-":
        return None
    if hasattr(v, "year") and hasattr(v, "month"):
        return int(datetime(v.year, v.month, v.day, tzinfo=timezone.utc).timestamp() * 1000)
    s = str(v).strip()
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return int(dt.timestamp() * 1000)
    except ValueError:
        return None


def _date_text(v) -> str:
    """离职日期(飞书文本型)→ YYYY-MM-DD。"""
    if v is None or v == "" or v == "-":
        return ""
    if hasattr(v, "year") and hasattr(v, "month"):
        return f"{v.year:04d}-{v.month:02d}-{v.day:02d}"
    s = str(v).strip()
    if len(s) >= 10 and s[4] in "-/.":
        return s[:10].replace("/", "-").replace(".", "-")
    return s


def _punch_ms(date_str: str, hhmm: str):
    """打卡时间(飞书日期型)→ 北京时间该日该时分的毫秒。"""
    if not hhmm or not date_str:
        return None
    try:
        hh, mm = hhmm.split(":")[:2]
        dt = datetime(int(date_str[:4]), int(date_str[5:7]), int(date_str[8:10]),
                      int(hh), int(mm), tzinfo=CN_TZ)
        return int(dt.timestamp() * 1000)
    except (ValueError, IndexError):
        return None


def parse_punch(v) -> tuple[str, str]:
    """打卡单元格 → (上班, 下班),取首末。"""
    if v is None:
        return "", ""
    s = str(v).strip()
    if not s or s == "-":
        return "", ""
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[-1]


# ---------------- 月份识别 ----------------
def detect_year_month(text: str) -> tuple[int, int]:
    """从文件名抓年月:'7月考勤报表_20260701_20260731' → (2026, 7)。"""
    m = re.search(r"(\d{4})(\d{2})(\d{2})", text)
    if m:
        mo = int(m.group(2))
        if 1 <= mo <= 12:
            return int(m.group(1)), mo
    m = re.search(r"(\d{1,2})月", text)
    if m:
        mo = int(m.group(1))
        if 1 <= mo <= 12:
            ym = re.search(r"(\d{4})年", text)
            return (int(ym.group(1)) if ym else 2026), mo
    raise RuntimeError(f"无法从文件名识别月份: {text}")


# ---------------- 解析原始报表(57列汇总 + 打卡明细) ----------------
def read_report(path: Path, year: int, month: int):
    """返回 (summary_records, detail_records)。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    row2 = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    summary_header = ["姓名"] + [
        str(row2[c - 1]).strip() if row2[c - 1] else "" for c in range(2, 58)
    ]
    month_prefix_date = f"{year:04d}-{month:02d}-"  # 如 2026-07-
    date_cols: list[tuple[int, str]] = []
    for c, v in enumerate(row2, 1):
        if not v:
            continue
        s = str(v).strip()
        # 打卡时间列:纯日期 "2026-07-01"(无星期)
        if len(s) == 10 and s.startswith(month_prefix_date) and "星期" not in s:
            date_cols.append((c, s))
        # 每日考勤结果列:"2026-07-01 星期三" —— 不灌,跳过

    summary: list[dict] = []
    detail: list[dict] = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or not str(name).strip():
            continue
        name = str(name).strip()
        gonghao = _text(ws.cell(row=r, column=2).value)

        rec: dict = {}
        for ci, h in enumerate(summary_header, 1):
            if not h:
                continue
            val = ws.cell(row=r, column=ci).value
            if h == "入职日期":
                ms = _date_ms_utc(val)
                if ms is not None:
                    rec[h] = ms
            elif h == "离职日期":
                t = _date_text(val)
                if t:
                    rec[h] = t
            elif h in SUMMARY_TEXT_COLS:
                t = _text(val)
                if t:
                    rec[h] = t
            else:
                n = _num(val)
                if n is not None:
                    rec[h] = n
        summary.append(rec)

        for col, dstr in date_cols:
            up, down = parse_punch(ws.cell(row=r, column=col).value)
            dms = _date_ms_utc(dstr)
            drec: dict = {"姓名": name}
            if gonghao:
                drec["工号"] = gonghao
            if dms is not None:
                drec["日期"] = dms
            up_ms = _punch_ms(dstr, up)
            if up_ms is not None:
                drec["上班打卡时间"] = up_ms
            down_ms = _punch_ms(dstr, down)
            if down_ms is not None:
                drec["下班打卡时间"] = down_ms
            detail.append(drec)
    wb.close()
    return summary, detail


# ---------------- 解析整理结果(3 个 sheet) ----------------
def _read_sheet_rows(ws, header_row: int = 3) -> list[dict]:
    """通用:表头在 header_row,数据从下一行起,到姓名空为止。"""
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        headers.append(str(v).strip() if v else "")
    rows: list[dict] = []
    r = header_row + 1
    while r <= ws.max_row:
        name = ws.cell(row=r, column=1).value
        if name is None or str(name).strip() == "":
            break  # 规则说明区/空行,停
        rec: dict = {}
        for ci, h in enumerate(headers, 1):
            if not h:
                continue
            rec[h] = ws.cell(row=r, column=ci).value
        rows.append(rec)
        r += 1
    return rows


def _conv_result(row: dict, person_map: dict[str, str] | None = None,
                 missing: set[str] | None = None) -> dict:
    """Sheet1 考勤整理结果 → 飞书16列。

    姓名/备注文本;姓名(人员)按 person_map 姓名→open_id 写 [{id:...}],不命中留空
    并把姓名记入 missing;是否确认复选不写(留空手勾);余数字。
    """
    out: dict = {}
    for k, v in row.items():
        if k in ("姓名", "备注"):
            t = _text(v)
            if t:
                out[k] = t
        else:
            n = _num(v)
            if n is not None:
                out[k] = n
    # 人员列:按姓名查 open_id 映射(飞书人员字段写 [{id: open_id}])
    name = out.get("姓名")
    if name and person_map is not None:
        oid = person_map.get(name)
        if oid:
            out["姓名 (人员 )"] = [{"id": oid}]
        elif missing is not None:
            missing.add(name)
    return out


def _conv_anomaly(row: dict) -> dict:
    """Sheet2 打卡异常说明 → 飞书6列,全文本(补卡次数/缺卡次数(上/下)原样)。"""
    out: dict = {}
    for k, v in row.items():
        t = _text(v)
        if t:
            out[k] = t
    return out


def _conv_other(row: dict) -> dict:
    """Sheet3 其他员工 → 飞书7列。姓名/人员类型/员工状态/备注文本,余数字。"""
    out: dict = {}
    for k, v in row.items():
        if k in ("姓名", "人员类型", "员工状态", "备注"):
            t = _text(v)
            if t:
                out[k] = t
        else:
            n = _num(v)
            if n is not None:
                out[k] = n
    return out


def read_result(path: Path, person_map: dict[str, str] | None = None):
    """返回 (result_rows, anomaly_rows, other_rows, missing_names)。

    result_rows 已按 person_map 填好姓名(人员)列;missing_names 是当月结果表
    里姓名不在总表映射中的人(需去飞书手补人员值)。
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    missing: set[str] = set()
    result = [_conv_result(r, person_map, missing)
              for r in _read_sheet_rows(wb["考勤整理结果"])]
    anomaly = [_conv_anomaly(r) for r in _read_sheet_rows(wb["打卡记录异常说明"])]
    other = [_conv_other(r) for r in _read_sheet_rows(wb["其他员工(无请假加班)"])]
    wb.close()
    return result, anomaly, other, sorted(missing)


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser(description="每月考勤上传飞书")
    ap.add_argument("report", help="考勤原始报表 xlsx 路径")
    ap.add_argument("result", help="考勤整理结果 xlsx 路径(含3 sheet)")
    ap.add_argument("--apply", action="store_true", help="真正建表+写入(否则仅预览)")
    args = ap.parse_args()

    report_path = Path(args.report)
    result_path = Path(args.result)
    if not report_path.exists():
        sys.exit(f"[错误] 报表不存在: {report_path}")
    if not result_path.exists():
        sys.exit(f"[错误] 整理结果不存在: {result_path}")

    year, month = detect_year_month(report_path.name)
    month_prefix = f"{month}月"
    print(f"[识别] {year}年{month}月 → 表前缀「{month_prefix}」")

    summary, detail = read_report(report_path, year, month)

    # 人员列需 open_id,从总表读 姓名→open_id 映射(只读 GET,干跑也读)
    token = get_token()
    person_map = load_person_map(token)

    result, anomaly, other, missing = read_result(result_path, person_map)

    # 月份列:考勤结果表每条记录带当月值,格式对齐总表(如 "2026-7",月不补零)
    month_value = f"{year}-{month}"
    for rec in result:
        rec["月份"] = month_value

    print(f"[解析] 汇总 {len(summary)} 行 | 打卡明细 {len(detail)} 行 | "
          f"结果 {len(result)} | 异常 {len(anomaly)} | 其他 {len(other)}")
    print(f"[月份] 考勤结果表每条写入「月份」= {month_value!r}")
    hit = len(result) - len(missing)
    print(f"[人员] 考勤结果表 {len(result)} 人,人员映射命中 {hit}, "
          f"未命中 {len(missing)}(将留空待手补)")
    if missing:
        print(f"[人员] 未命中姓名: {'、'.join(missing)}")

    if not args.apply:
        print("\n[干跑] 未加 --apply,不建表不写入。各表样本:")
        print(f"  汇总[0]: {summary[0] if summary else '(空)'}")
        print(f"  结果[0]: {result[0] if result else '(空)'}")
        print(f"  异常[0]: {anomaly[0] if anomaly else '(空)'}")
        print(f"  其他[0]: {other[0] if other else '(空)'}")
        print(f"\n将建/复用 5 张表:")
        for key in TABLE_SPECS:
            print(f"  「{month_prefix}{key}」{len(TABLE_SPECS[key])}字段")
        print("\n确认无误后加 --apply 真正执行。")
        return

    existing = list_tables(token)
    print(f"\n[建表] base 现有 {len(existing)} 张表")
    tids: dict[str, str] = {}
    for key, fields in TABLE_SPECS.items():
        tids[key] = ensure_table(token, month_prefix, key, fields, existing)

    # 写入(月表语义=当月最新,先清空再写)
    plan = [
        ("考勤报表", summary, "汇总"),
        ("打卡时间明细", detail, "明细"),
        ("考勤结果表", result, "结果"),
        ("打卡异常说明", anomaly, "异常"),
        ("其他员工", other, "其他"),
    ]
    print("\n[写入]")
    counts = {}
    for key, records, label in plan:
        tid = tids[key]
        print(f"  「{month_prefix}{key}」({label}) {len(records)} 条")
        ids = list_record_ids(token, tid)
        if ids:
            print(f"    清空旧 record {len(ids)} 条")
            batch_delete(token, tid, ids)
        if records:
            counts[key] = batch_create(token, tid, records)
        else:
            counts[key] = 0
            print(f"    (无数据,跳过写入)")

    print(f"\n✅ 完成「{month_prefix}」:")
    for key, label in [("考勤报表","汇总"),("打卡时间明细","明细"),("考勤结果表","结果"),
                       ("打卡异常说明","异常"),("其他员工","其他")]:
        print(f"   {label}: {counts.get(key,0)} 条")

    # 覆盖式写入考勤结果总表(清空旧数据 → 全量写入当月)
    master_tid = existing.get(PERSON_MAP_TABLE)
    if master_tid and result:
        print(f"\n[总表] 覆盖写入「{PERSON_MAP_TABLE}」({month_value})")
        cr = sync_master(token, master_tid, result, month_value)
        print(f"  写入 {cr} 条")
    elif not master_tid:
        print(f"\n[总表] 未找到「{PERSON_MAP_TABLE}」,跳过。请在飞书手建一张同名字段表。")

    if missing:
        print(f"\n⚠️ 「{month_prefix}考勤结果表」以下 {len(missing)} 人的人员列留空"
              f"(总表无其姓名→人员映射),请去飞书手补:")
        print(f"   {'、'.join(missing)}")


if __name__ == "__main__":
    main()
