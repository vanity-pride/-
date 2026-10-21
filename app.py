# -*- coding: utf-8 -*-
"""
每日打卡记录系统（Flask + openpyxl）

功能：
  1. 多级分类结构：分类(group) 可嵌套子分类 / 勾选项(check) / 数值项(number) / 多次打卡项(multi)
  2. 每日勾选打卡：已完成 = 1，未完成 = 0
  3. 多次打卡项：可设「目标次数」（如八部金刚功每日 4 次），当日记录次数 0~目标，
     次数 ≥ 目标才算当天完成，收益按「达标才结算一次固定额」计
  4. 收益属性：勾选项/多次项可设定收益金额（如八部金刚功 = 0.5 元）；
     数值项直接以填入数值作为收益，支持负数，无单位
  5. 每日收益表：当日新增 = 各勾选项(勾选×收益) + 各多次项(达标×收益) + 各数值项当日数值；
     累计收益 = 上一日累计 + 当日新增，初始值 -8407
  6. 历史补录：可给过去某天补勾选/补次数/补数值，收益链自动重算
  7. 数据分析：365 天热力图、近 12 周/12 月趋势、分类贡献、收益目标线
  8. 数据安全：保存前自动备份（每天一份，保留最近 30 份）、临时文件原子写入、
     Excel 占用检测与友好提示
  9. Excel 持久化：「活动结构」「打卡记录」「每日收益」三个 sheet
  10. 已过去日期且未填写的项目自动置 0；历史记录 + 当日状态 + 收益曲线

运行：conda run -n work python app.py → http://127.0.0.1:5000
"""
import glob
import io
import json
import os
import shutil
import threading
import zipfile
from calendar import monthrange
from datetime import date, datetime, timedelta

from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, "data", "打卡记录.xlsx")
SHEET_RECORD = "打卡记录"
SHEET_STRUCT = "活动结构"
SHEET_BENEFIT = "每日收益"
SHEET_ARCHIVE = "归档记录"   # 冷表：已归档项目的历史数据（项目ID / 名称 / 日期 / 值）
STRUCT_HEADERS = ["ID", "名称", "类型", "父级ID", "单位", "收益", "目标次数", "状态"]
BENEFIT_HEADERS = ["日期", "当日新增", "累计收益"]
ARCHIVE_HEADERS = ["项目ID", "项目名称", "日期", "值"]
INITIAL_BENEFIT = -8407.0  # 初始收益（此前手动记录的数据）

WEEKDAYS_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
MAX_NAME_LEN = 20
MAX_UNIT_LEN = 6
MAX_AMOUNT = 1e9
MAX_TARGET = 99           # 多次打卡项目标次数上限
MAX_BACKUPS = 30          # 自动备份保留份数（每天一份）
MAX_CHANGES = 200         # 单次补录最多修改的项目数
HEATMAP_DAYS = 365        # 热力图覆盖天数
TREND_WEEKS = 12          # 周趋势档期数
TREND_MONTHS = 12         # 月趋势档期数

TYPE_GROUP = "group"    # 分类
TYPE_CHECK = "check"    # 勾选项（1/0），带收益属性
TYPE_NUMBER = "number"  # 数值项，数值本身即收益
TYPE_MULTI = "multi"    # 多次打卡项（当日次数 0~目标次数，达标才结算收益）

STATUS_ACTIVE = "active"      # 正常项目中
STATUS_ARCHIVED = "archived"  # 已归档（软删除）：历史数据完整保留，仅退出日常打卡

# 首次初始化时的示例结构：(名称, 类型, 父级名, 单位, 收益, 目标次数)
DEMO_STRUCTURE = [
    ("养生", TYPE_GROUP, None, "", 0, 1),
    ("八部金刚功", TYPE_MULTI, "养生", "", 0.5, 4),
    ("肾俞", TYPE_CHECK, "养生", "", 0.5, 1),
    ("专注时间", TYPE_NUMBER, None, "", 0, 1),
]

# 旧版升级时自动识别为多次项的勾选项名称 -> 目标次数（历史打勾值 1 视为已达标）
LEGACY_MULTI_UPGRADE = {"八部金刚功": 4, "八步金刚功": 4}

app = Flask(__name__)
_lock = threading.RLock()


# ---------------- 路径 / 设置 ----------------

def _data_dir():
    return os.path.dirname(XLSX_PATH)


def _backup_dir():
    return os.path.join(_data_dir(), "backup")


def _settings_path():
    return os.path.join(_data_dir(), "settings.json")


def _load_settings():
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            s = json.load(f)
    except (OSError, ValueError):
        s = {}
    if not isinstance(s, dict):
        s = {}
    try:
        goal = float(s.get("goal", 0) or 0)
    except (TypeError, ValueError):
        goal = 0.0
    if goal != goal or abs(goal) >= MAX_AMOUNT:
        goal = 0.0
    return {"goal": round(goal, 2), "note": str(s.get("note", "") or "")[:30]}


def _save_settings(s):
    os.makedirs(_data_dir(), exist_ok=True)
    tmp = _settings_path() + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"goal": s.get("goal", 0), "note": str(s.get("note", "") or "")[:30]},
                  f, ensure_ascii=False, indent=2)
    os.replace(tmp, _settings_path())


# ---------------- 文件 / 工作簿 ----------------

def _create_workbook():
    """首次运行：创建示例结构。"""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_RECORD
    ws.cell(row=1, column=1, value="日期")
    wss = wb.create_sheet(SHEET_STRUCT)
    for i, h in enumerate(STRUCT_HEADERS, start=1):
        wss.cell(row=1, column=i, value=h)
    name_to_id = {}
    nid = 1
    for name, typ, parent, unit, benefit, target in DEMO_STRUCTURE:
        r = nid + 1
        wss.cell(row=r, column=1, value=nid)
        wss.cell(row=r, column=2, value=name)
        wss.cell(row=r, column=3, value=typ)
        wss.cell(row=r, column=4, value=name_to_id.get(parent) if parent else None)
        wss.cell(row=r, column=5, value=unit)
        wss.cell(row=r, column=6, value=benefit)
        wss.cell(row=r, column=7, value=target)
        wss.cell(row=r, column=8, value=STATUS_ACTIVE)
        name_to_id[name] = nid
        nid += 1
        if typ != TYPE_GROUP:
            ws.cell(row=1, column=ws.max_column + 1, value=name)
    return wb


def _migrate_to_structure(wb):
    """旧版扁平活动列表 → 新版多级结构（保留全部历史数据，收益默认 0）。"""
    wsr = wb[SHEET_RECORD]
    cols = []
    for col in range(2, wsr.max_column + 1):
        v = wsr.cell(row=1, column=col).value
        if v not in (None, ""):
            cols.append(str(v).strip())
    wss = wb.create_sheet(SHEET_STRUCT)
    for i, h in enumerate(STRUCT_HEADERS, start=1):
        wss.cell(row=1, column=i, value=h)
    nid = 1
    for name in cols:
        r = nid + 1
        wss.cell(row=r, column=1, value=nid)
        wss.cell(row=r, column=2, value=name)
        wss.cell(row=r, column=3, value=TYPE_CHECK)
        wss.cell(row=r, column=4, value=None)
        wss.cell(row=r, column=5, value="")
        wss.cell(row=r, column=6, value=0)
        wss.cell(row=r, column=7, value=1)
        wss.cell(row=r, column=8, value=STATUS_ACTIVE)
        nid += 1
    _save(wb)


def _migrate_multi(wb):
    """
    多次打卡项迁移（幂等，可重复执行）：
      1. 结构表缺「目标次数」列时补上，默认 1；
      2. 仅在首次遇到缺列的旧版文件时，把其中的金刚功（LEGACY_MULTI_UPGRADE）
         自动升级为多次项，目标次数按映射表设定，打卡记录中该列历史值 1 → 目标次数
         （视为达标），保证历史收益与完成率不变。
         注意：升级是一次性的——之后用户把多次项转回勾选项，不会被这里强制升级回去。
    """
    wss = wb[SHEET_STRUCT]
    col_t = len(STRUCT_HEADERS)  # 「目标次数」列号
    found = False
    for c in range(1, wss.max_column + 1):
        if str(wss.cell(row=1, column=c).value or "").strip() == "目标次数":
            col_t = c
            found = True
            break
    if not found:
        for c in range(1, wss.max_column + 2):
            if wss.cell(row=1, column=c).value in (None, ""):
                col_t = c
                break
        wss.cell(row=1, column=col_t, value="目标次数")

    upgraded = {}  # 名称 -> 目标次数
    need_default = False
    for r in range(2, wss.max_row + 1):
        vid = wss.cell(row=r, column=1).value
        if vid in (None, ""):
            continue
        name = str(wss.cell(row=r, column=2).value or "").strip()
        typ = str(wss.cell(row=r, column=3).value or "").strip() or TYPE_CHECK
        # LEGACY 升级只在「目标次数」列刚补上时执行（首次遇到旧版文件）。
        # 否则用户把多次项转回勾选项后，每次加载都会被强制升级回去。
        if not found and typ == TYPE_CHECK and name in LEGACY_MULTI_UPGRADE:
            tgt = LEGACY_MULTI_UPGRADE[name]
            wss.cell(row=r, column=3, value=TYPE_MULTI)
            wss.cell(row=r, column=col_t, value=tgt)
            upgraded[name] = tgt
        elif wss.cell(row=r, column=col_t).value in (None, ""):
            wss.cell(row=r, column=col_t, value=1)
            need_default = True

    record_changed = False
    if upgraded and SHEET_RECORD in wb.sheetnames:
        wsr = wb[SHEET_RECORD]
        for col in range(2, wsr.max_column + 1):
            hdr = wsr.cell(row=1, column=col).value
            if hdr in upgraded:
                for r in range(2, wsr.max_row + 1):
                    cell = wsr.cell(row=r, column=col)
                    if cell.value in (1, "1"):
                        cell.value = upgraded[hdr]
                        record_changed = True

    if not found or need_default or upgraded or record_changed:
        _save(wb)


def _migrate_status(wb):
    """状态列迁移（幂等）：结构表缺「状态」列时补上，默认 active。"""
    wss = wb[SHEET_STRUCT]
    for c in range(1, wss.max_column + 1):
        if str(wss.cell(row=1, column=c).value or "").strip() == "状态":
            return
    col_s = wss.max_column + 1
    wss.cell(row=1, column=col_s, value="状态")
    for r in range(2, wss.max_row + 1):
        if wss.cell(row=r, column=1).value not in (None, ""):
            wss.cell(row=r, column=col_s, value=STATUS_ACTIVE)
    _save(wb)


def _load():
    path = XLSX_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if not os.path.exists(path):
        wb = _create_workbook()
        _save(wb)
    wb = load_workbook(path)
    if SHEET_STRUCT not in wb.sheetnames:
        _migrate_to_structure(wb)
        wb = load_workbook(path)
    if SHEET_RECORD not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_RECORD, 0)
        ws.cell(row=1, column=1, value="日期")
        _save(wb)
    _migrate_multi(wb)
    _migrate_status(wb)
    return wb


def _owner_file():
    """Excel 打开工作簿时生成的属主锁文件（~$xxx.xlsx）。"""
    return os.path.join(_data_dir(), "~$" + os.path.basename(XLSX_PATH))


def _check_writable():
    if os.path.exists(_owner_file()):
        raise PermissionError("数据文件正在 Excel 中打开，请先关闭 data/打卡记录.xlsx 再操作")


def _backup(force=False):
    """备份数据文件：默认每天一份（保留最近 MAX_BACKUPS 份）；force 为手动备份（带时间戳）。"""
    if not os.path.exists(XLSX_PATH):
        return None
    try:
        os.makedirs(_backup_dir(), exist_ok=True)
        if force:
            dst = os.path.join(_backup_dir(),
                               "打卡记录_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S"))
        else:
            dst = os.path.join(_backup_dir(), "打卡记录_%s.xlsx" % date.today().isoformat())
            if os.path.exists(dst):
                return dst
        shutil.copy2(XLSX_PATH, dst)
        files = sorted(glob.glob(os.path.join(_backup_dir(), "打卡记录_*.xlsx")))
        for f in files[:-MAX_BACKUPS]:
            try:
                os.remove(f)
            except OSError:
                pass
        return dst
    except OSError:
        return None


def _list_backups():
    out = []
    for f in sorted(glob.glob(os.path.join(_backup_dir(), "打卡记录_*.xlsx")), reverse=True):
        try:
            st = os.stat(f)
            out.append({"name": os.path.basename(f), "size": st.st_size,
                        "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")})
        except OSError:
            continue
    return out


def _save(wb):
    """先备份、再写临时文件、最后原子替换，避免写入中途出错损坏数据文件。"""
    _check_writable()
    _backup()
    tmp = XLSX_PATH + ".tmp"
    try:
        wb.save(tmp)
        os.replace(tmp, XLSX_PATH)
    except PermissionError:
        raise PermissionError("数据文件正在 Excel 中打开，请先关闭 data/打卡记录.xlsx 再操作")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def _sheet(wb):
    return wb[SHEET_RECORD]


# ---------------- 结构（活动结构表） ----------------

def _read_structure(wb):
    ws = wb[SHEET_STRUCT]
    # 「目标次数」「状态」列位置（兼容旧文件缺列的情况）
    col_t = len(STRUCT_HEADERS)
    col_s = len(STRUCT_HEADERS)
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=1, column=c).value or "").strip()
        if h == "目标次数":
            col_t = c
        elif h == "状态":
            col_s = c
    items = []
    for r in range(2, ws.max_row + 1):
        vid = ws.cell(row=r, column=1).value
        if vid in (None, ""):
            continue
        try:
            vid = int(vid)
        except (TypeError, ValueError):
            continue
        name = str(ws.cell(row=r, column=2).value or "").strip()
        typ = str(ws.cell(row=r, column=3).value or "").strip() or TYPE_CHECK
        parent = ws.cell(row=r, column=4).value
        try:
            parent = int(parent) if parent not in (None, "") else None
        except (TypeError, ValueError):
            parent = None
        unit = str(ws.cell(row=r, column=5).value or "").strip()
        try:
            benefit = float(ws.cell(row=r, column=6).value or 0)
        except (TypeError, ValueError):
            benefit = 0.0
        target = _parse_target(ws.cell(row=r, column=col_t).value)
        if typ != TYPE_MULTI:
            target = 1  # 非多次项目标恒为 1
        status = str(ws.cell(row=r, column=col_s).value or "").strip() or STATUS_ACTIVE
        if status != STATUS_ARCHIVED:
            status = STATUS_ACTIVE
        items.append({"id": vid, "name": name, "type": typ, "parent": parent,
                      "unit": unit, "benefit": round(benefit, 2),
                      "target": target, "status": status})
    return items


def _write_structure(wb, items):
    if SHEET_STRUCT in wb.sheetnames:
        del wb[SHEET_STRUCT]
    ws = wb.create_sheet(SHEET_STRUCT)
    for i, h in enumerate(STRUCT_HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    for idx, it in enumerate(items, start=2):
        ws.cell(row=idx, column=1, value=it["id"])
        ws.cell(row=idx, column=2, value=it["name"])
        ws.cell(row=idx, column=3, value=it["type"])
        ws.cell(row=idx, column=4, value=it["parent"])
        ws.cell(row=idx, column=5, value="")  # 单位字段已废弃，保留列兼容旧文件
        ws.cell(row=idx, column=6, value=it.get("benefit", 0))
        ws.cell(row=idx, column=7, value=it.get("target", 1))
        ws.cell(row=idx, column=8, value=it.get("status", STATUS_ACTIVE))
    return ws


def _children_map(items):
    known = {it["id"] for it in items}
    ch = {}
    for it in items:
        it["parent"] = it["parent"] if it["parent"] in known else None  # 孤儿归一到顶级
        ch.setdefault(it["parent"], []).append(it)
    return ch


def _leaf_order(items):
    """按层级 DFS 输出所有叶子项（勾选项 + 数值项）的顺序。"""
    ch = _children_map(items)
    out = []

    def dfs(node):
        if node["type"] != TYPE_GROUP:
            out.append(node)
        for c in ch.get(node["id"], []):
            dfs(c)

    for root in ch.get(None, []):
        dfs(root)
    return out


def _build_tree(items, archived=False):
    """
    构建项目树。archived=False 只含正常项目（今日打卡 / 日常统计用）；
    archived=True 只含归档项目（归档箱用）。
    归档视图中，父级仍正常的归档项会作为独立子树根展示，保证不漏项。
    节点带 status 字段。
    """
    ch = _children_map(items)

    def visible(node):
        return (node.get("status") == STATUS_ARCHIVED) if archived \
            else (node.get("status") != STATUS_ARCHIVED)

    def archive_roots(node):
        """node（自身未归档）之下处于归档链顶层的子项。"""
        res = []
        for c in ch.get(node["id"], []):
            if visible(c):
                res.append(c)
            else:
                res.extend(archive_roots(c))
        return res

    roots = []
    for r in ch.get(None, []):
        if visible(r):
            roots.append(r)
        elif archived:
            roots.extend(archive_roots(r))

    def build(node):
        n = {"id": node["id"], "name": node["name"], "type": node["type"],
             "parent": node["parent"], "benefit": node.get("benefit", 0),
             "target": node.get("target", 1), "status": node.get("status", STATUS_ACTIVE)}
        if node["type"] == TYPE_GROUP:
            n["children"] = [build(c) for c in ch.get(node["id"], []) if visible(c)]
        return n

    return [build(r) for r in roots]


def _paths(items):
    by_id = {it["id"]: it for it in items}

    def path(it):
        parts = [it["name"]]
        p, guard = it["parent"], 0
        while p is not None and p in by_id and guard < 20:
            parts.append(by_id[p]["name"])
            p = by_id[p]["parent"]
            guard += 1
        return " / ".join(reversed(parts))

    return {it["id"]: path(it) for it in items}


def _find(items, item_id):
    for it in items:
        if it["id"] == item_id:
            return it
    return None


def _validate_name(items, name, exclude_id=None):
    if not name:
        raise ValueError("名称不能为空")
    if len(name) > MAX_NAME_LEN:
        raise ValueError("名称不能超过 %d 个字符" % MAX_NAME_LEN)
    dup = next((it for it in items
                if it["name"] == name and it["id"] != exclude_id), None)
    if dup is not None:
        if dup.get("status") == STATUS_ARCHIVED:
            raise ValueError(
                "名称「%s」被已归档项目占用。解决办法：到「结构管理 → 归档箱」"
                "把该项重命名为历史版本（如「%s·旧」），即可释放该名称。"
                % (name, name))
        raise ValueError("名称已存在")


def _validate_parent(items, parent_id, self_id=None):
    if parent_id is None:
        return
    p = _find(items, parent_id)
    if p is None or p["type"] != TYPE_GROUP:
        raise ValueError("父级必须是一个分类")
    if p.get("status") == STATUS_ARCHIVED:
        raise ValueError("父级已归档，请先恢复该分类")
    cur = parent_id
    while cur is not None:
        if cur == self_id:
            raise ValueError("不能将分类移动到它自己的子级下")
        nxt = _find(items, cur)
        cur = nxt["parent"] if nxt else None


def _parse_amount(v, label="收益"):
    try:
        x = float(v)
    except (TypeError, ValueError):
        raise ValueError("%s数值无效" % label)
    if x != x or abs(x) >= MAX_AMOUNT:  # NaN 或超范围
        raise ValueError("%s数值超出范围" % label)
    return round(x, 2)


def _parse_target(v):
    """目标次数：1~MAX_TARGET 的整数，非法值回落为 1。"""
    try:
        t = int(float(v))
    except (TypeError, ValueError):
        return 1
    if t < 1:
        return 1
    return min(t, MAX_TARGET)


def _add_item(items, name, typ, parent_id, benefit=0, target=1):
    _validate_name(items, name)
    if typ not in (TYPE_GROUP, TYPE_CHECK, TYPE_NUMBER, TYPE_MULTI):
        raise ValueError("类型不正确")
    _validate_parent(items, parent_id)
    if typ in (TYPE_CHECK, TYPE_MULTI):
        benefit = _parse_amount(benefit)
    else:
        benefit = 0  # 数值项收益=填入数值；分类无收益
    if typ == TYPE_MULTI:
        target = _parse_target(target)
    else:
        target = 1
    new_id = max((it["id"] for it in items), default=0) + 1
    items.append({"id": new_id, "name": name, "type": typ, "parent": parent_id,
                  "unit": "", "benefit": benefit, "target": target,
                  "status": STATUS_ACTIVE})
    return new_id


def _update_item(items, item_id, name, parent_id, benefit=None, target=None, type_=None):
    it = _find(items, item_id)
    if it is None:
        raise KeyError("项目不存在")
    it.setdefault("status", STATUS_ACTIVE)
    if it["status"] == STATUS_ARCHIVED:
        raise ValueError("归档项不可修改属性，只能重命名（见归档箱）")
    _validate_name(items, name, exclude_id=item_id)
    _validate_parent(items, parent_id, self_id=item_id)
    old_type = it["type"]
    if type_ and type_ != old_type:
        # 只允许勾选项 <-> 多次项互转：两者都是"完成型"计数语义
        if type_ not in (TYPE_CHECK, TYPE_MULTI) or old_type not in (TYPE_CHECK, TYPE_MULTI):
            raise ValueError("仅支持「勾选项」与「多次打卡项」之间互相转换")
        it["type"] = type_
    if it["type"] in (TYPE_CHECK, TYPE_MULTI):
        it["benefit"] = _parse_amount(benefit) if benefit is not None else 0
    else:
        it["benefit"] = 0
    if it["type"] == TYPE_MULTI:
        it["target"] = _parse_target(target) if target is not None else 1
    else:
        it["target"] = 1
    it["name"] = name
    it["parent"] = parent_id
    return old_type


def _convert_type_records(ws, items, item_id, old_type, new_type):
    """类型互转时改写历史列的值，保持「当天做过」的语义：
    - 多次项 -> 勾选项：练过（次数 > 0）即算勾选，记 1
    - 勾选项 -> 多次项：0/1 原样保留（1 次即达标，收益不变）
    """
    if old_type == new_type:
        return
    it = _find(items, item_id)
    if it is None or it["type"] == TYPE_GROUP:
        return
    col = _name_to_col(ws).get(it["name"])
    if col is None:
        return
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=col)
        v = cell.value
        if v in (None, ""):
            continue
        if old_type == TYPE_MULTI and new_type == TYPE_CHECK:
            cell.value = 1 if _num(v) > 0 else 0
        # check -> multi：值域 0/1 本身合法，无需改写


def _orphan_columns(ws, items):
    """记录表中结构里不存在的列（通常是手工改了「活动结构」表的名称/删项导致）。
    返回孤儿列名列表；正常应为空。"""
    leaves = {l["name"] for l in _leaf_order(items)}
    return [str(v).strip() for v in (_name_to_col(ws) or {}).keys() if str(v).strip() not in leaves]


def _rename_item(items, item_id, name):
    """重命名（支持归档项）：只改名称，其余属性（类型/父级/收益/状态）不动。
    调用方负责同步改写记录表的列名（_rewrite_records 会按名称搬移历史数据）。"""
    it = _find(items, item_id)
    if it is None:
        raise KeyError("项目不存在")
    _validate_name(items, name, exclude_id=item_id)
    it["name"] = name


def _subtree_ids(items, item_id):
    """返回以 item_id 为根的子树全部 id（含自身）。"""
    ids = {item_id}
    changed = True
    while changed:
        changed = False
        for it in items:
            if it["parent"] in ids and it["id"] not in ids:
                ids.add(it["id"])
                changed = True
    return ids


def _ancestor_ids(items, item_id):
    """返回 item_id 的全部祖先 id（不含自身）。"""
    by_id = {it["id"]: it for it in items}
    out, cur, guard = set(), by_id.get(item_id), 0
    while cur is not None and cur.get("parent") is not None and guard < 50:
        out.add(cur["parent"])
        cur = by_id.get(cur["parent"])
        guard += 1
    return out


def _set_status(items, ids, status):
    for it in items:
        if it["id"] in ids:
            it["status"] = status


def _archive_subtree(items, item_id):
    """软删除：把整棵子树标记为归档。历史记录列与收益链完全不动。"""
    if _find(items, item_id) is None:
        raise KeyError("项目不存在")
    ids = _subtree_ids(items, item_id)
    _set_status(items, ids, STATUS_ARCHIVED)
    return ids


def _restore_subtree(items, item_id):
    """
    恢复归档项：连同其祖先（向上闭合，保证 active 树结构完整）
    与全部后代一起恢复为正常状态。
    """
    if _find(items, item_id) is None:
        raise KeyError("项目不存在")
    ids = _subtree_ids(items, item_id) | _ancestor_ids(items, item_id)
    _set_status(items, ids, STATUS_ACTIVE)
    return ids


def _purge_subtree(items, item_id):
    """永久删除：从结构表移除整棵子树（调用方负责删记录列并重算收益链）。"""
    if _find(items, item_id) is None:
        raise KeyError("项目不存在")
    ids = _subtree_ids(items, item_id)
    kept = [it for it in items if it["id"] not in ids]
    return kept, ids


# ---------------- 记录（打卡记录表） ----------------

def _read_records(ws):
    """日期字符串 -> {列名: 原始值}，以及表头列名顺序。"""
    names = []
    for col in range(2, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v not in (None, ""):
            names.append(str(v).strip())
    records = {}
    for r in range(2, ws.max_row + 1):
        d = ws.cell(row=r, column=1).value
        if d in (None, ""):
            continue
        vals = {}
        for i, col in enumerate(range(2, 2 + len(names))):
            vals[names[i]] = ws.cell(row=r, column=col).value
        records[str(d).strip()] = vals
    return records, names


def _name_to_col(ws):
    m = {}
    for col in range(2, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v not in (None, ""):
            m[str(v).strip()] = col
    return m


# ---------------- 归档记录（冷表：已归档项目的历史数据） ----------------

def _cold_blank(v):
    """冷表只存非零值：0 / 空 在勾选/多次/数值三类项上都等价于「当天无贡献」。"""
    return v in (None, "") or _num(v) == 0


def _read_cold(wb):
    """归档记录表 -> {项目ID: {"name": 名称快照, "days": {日期: 值}}}"""
    out = {}
    if SHEET_ARCHIVE not in wb.sheetnames:
        return out
    ws = wb[SHEET_ARCHIVE]
    for r in range(2, ws.max_row + 1):
        iid = ws.cell(row=r, column=1).value
        if iid in (None, ""):
            continue
        d = str(ws.cell(row=r, column=3).value or "").strip()
        if not d:
            continue
        ent = out.setdefault(int(iid), {"name": "", "days": {}})
        nm = str(ws.cell(row=r, column=2).value or "").strip()
        if nm:
            ent["name"] = nm
        v = ws.cell(row=r, column=4).value
        if v not in (None, ""):
            ent["days"][d] = v
    return out


def _write_cold(wb, cold):
    """按 {项目ID: {"name", "days"}} 全量重写归档记录表（数据量小，全量重写最稳）。"""
    if SHEET_ARCHIVE in wb.sheetnames:
        del wb[SHEET_ARCHIVE]
    ws = wb.create_sheet(SHEET_ARCHIVE)
    for i, h in enumerate(ARCHIVE_HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    r = 2
    for iid in sorted(cold):
        ent = cold[iid]
        for d in sorted(ent["days"]):
            ws.cell(row=r, column=1, value=iid)
            ws.cell(row=r, column=2, value=ent["name"])
            ws.cell(row=r, column=3, value=d)
            ws.cell(row=r, column=4, value=ent["days"][d])
            r += 1
    return ws


def _migrate_legacy_archived(wb, items):
    """幂等迁移：结构里已归档、但主表仍占着列的项目 -> 整列搬入冷表。返回是否有搬迁。"""
    if SHEET_RECORD not in wb.sheetnames:
        return False
    ws = wb[SHEET_RECORD]
    records, _ = _read_records(ws)
    n2c = _name_to_col(ws)
    cold = _read_cold(wb)
    moved = False
    for leaf in _leaf_order(items):
        if leaf.get("status") != STATUS_ARCHIVED or leaf["name"] not in n2c:
            continue
        if leaf["id"] in cold:          # 已迁过，保持幂等
            continue
        days = {d: vals[leaf["name"]] for d, vals in records.items()
                if not _cold_blank(vals.get(leaf["name"]))}
        cold[leaf["id"]] = {"name": leaf["name"], "days": days}
        moved = True
    if moved:
        _write_cold(wb, cold)
    return moved


def _migrate_to_cold(wb, items, ids):
    """把指定项目（须为已归档叶子项）的主表列数据搬入冷表。返回冷表 dict。"""
    ws = wb[SHEET_RECORD]
    records, _ = _read_records(ws)
    cold = _read_cold(wb)
    by_id = {it["id"]: it for it in items}
    for iid in ids:
        it = by_id.get(iid)
        if it is None or it.get("status") != STATUS_ARCHIVED or it["type"] == TYPE_GROUP:
            continue
        if iid in cold:
            continue
        days = {d: vals[it["name"]] for d, vals in records.items()
                if not _cold_blank(vals.get(it["name"]))}
        cold[iid] = {"name": it["name"], "days": days}
    _write_cold(wb, cold)
    return cold


def _restore_from_cold(wb, items, ids):
    """恢复归档项：把冷表中的历史值写回主表列（列须已存在），并删除冷表对应行。"""
    ws = wb[SHEET_RECORD]
    n2c = _name_to_col(ws)
    cold = _read_cold(wb)
    by_id = {it["id"]: it for it in items}
    for iid in ids:
        ent = cold.get(iid)
        it = by_id.get(iid)
        if ent is None or it is None or it.get("status") == STATUS_ARCHIVED:
            continue
        col = n2c.get(it["name"])
        if col is None:
            continue
        for d, v in sorted(ent["days"].items()):
            ws.cell(row=_ensure_row(ws, d), column=col).value = v
        cold.pop(iid)
    _write_cold(wb, cold)


def _purge_cold(wb, ids):
    """永久删除：移除冷表中这些项目的历史行。"""
    cold = _read_cold(wb)
    keep = {k: v for k, v in cold.items() if k not in ids}
    if len(keep) != len(cold):
        _write_cold(wb, keep)


def _sync_cold_name(wb, item_id, new_name):
    """重命名归档项时同步冷表中的名称快照（保持人读一致）。"""
    cold = _read_cold(wb)
    ent = cold.get(item_id)
    if ent is not None and ent["name"] != new_name:
        ent["name"] = new_name
        _write_cold(wb, cold)


def _read_all_records(wb, items):
    """主表 + 冷表合并视图：{日期: {项目名: 值}}（归档项按当前名称注入，下游无感）。"""
    records, names = _read_records(wb[SHEET_RECORD])
    by_id = {it["id"]: it for it in items}
    for iid, ent in _read_cold(wb).items():
        it = by_id.get(iid)
        if it is None:
            continue
        for d, v in ent["days"].items():
            records.setdefault(d, {})[it["name"]] = v
    return records, names


def _ensure_row(ws, ds):
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v not in (None, "") and str(v).strip() == ds:
            return r
    r = ws.max_row + 1
    ws.cell(row=r, column=1, value=ds)
    return r


def _sort_rows(ws):
    """按日期升序重排数据行（补录过去日期后调用，保持 Excel 里行序整洁）。"""
    maxr, maxc = ws.max_row, ws.max_column
    if maxr < 2:
        return
    data = []
    for r in range(2, maxr + 1):
        d = ws.cell(row=r, column=1).value
        if d in (None, ""):
            continue
        data.append((str(d).strip(),
                     [ws.cell(row=r, column=c).value for c in range(2, maxc + 1)]))
    if not data:
        return
    data.sort(key=lambda x: x[0])
    ws.delete_rows(2, maxr - 1)
    for i, (d, vals) in enumerate(data, start=2):
        ws.cell(row=i, column=1, value=d)
        for j, v in enumerate(vals, start=2):
            if v not in (None, ""):
                ws.cell(row=i, column=j, value=v)


def _sync_records(ws, items):
    """
    补日期行、补缺失列，并把"过去（含今天）未填写"的单元格一律写 0。
    规则：若当日已过且该项未被勾选/未填数值，系统自动将其设为 0；
    多次打卡项的次数钳制到 0~目标次数。
    归档项的历史在冷表，不参与主表补列与补 0。
    """
    leaves = [l for l in _leaf_order(items) if l.get("status") != STATUS_ARCHIVED]
    if not leaves:
        return False
    n2c = _name_to_col(ws)
    max_col = ws.max_column
    changed = False
    for leaf in leaves:  # 结构里有但表里缺的列，补上
        if leaf["name"] not in n2c:
            max_col += 1
            ws.cell(row=1, column=max_col, value=leaf["name"])
            n2c[leaf["name"]] = max_col
            changed = True

    today = date.today()
    idx = {}
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v not in (None, ""):
            idx[str(v).strip()] = r
    start = min((datetime.strptime(d, "%Y-%m-%d").date() for d in idx), default=today)

    next_row = ws.max_row + 1
    cur = start
    while cur <= today:
        ds = cur.isoformat()
        if ds not in idx:
            ws.cell(row=next_row, column=1, value=ds)
            idx[ds] = next_row
            next_row += 1
            changed = True
        r = idx[ds]
        for leaf in leaves:
            cell = ws.cell(row=r, column=n2c[leaf["name"]])
            if cell.value in (None, ""):
                cell.value = 0
                changed = True
            elif leaf["type"] == TYPE_MULTI:
                clamped = min(max(int(_num(cell.value)), 0), leaf.get("target", 1))
                if clamped != cell.value:
                    cell.value = clamped
                    changed = True
        cur += timedelta(days=1)
    return changed


def _rename_map(old_items, new_items):
    """对比改名前后的项目表，产出 {旧名: 新名}（仅名称发生变化的项）."""
    old_by_id = {it["id"]: it["name"] for it in old_items}
    out = {}
    for it in new_items:
        old = old_by_id.get(it["id"])
        if old is not None and old != it["name"]:
            out[old] = it["name"]
    return out


def _rewrite_records(wb, items, renames=None):
    """按层级 DFS 顺序重排记录表列，保留全部历史数值。
    renames: {旧名: 新名}，重命名后旧列的数据按映射归位到新列名。
    归档项的列已迁入冷表，这里只重排在役项目的列。"""
    _migrate_legacy_archived(wb, items)   # 兜底：旧文件的归档列先搬冷表，避免被丢弃
    wsr = wb[SHEET_RECORD]
    records, _ = _read_records(wsr)
    if renames:
        records = {d: {renames.get(k, k): v for k, v in vals.items()}
                   for d, vals in records.items()}
    leaves = [l for l in _leaf_order(items) if l.get("status") != STATUS_ARCHIVED]
    names = [l["name"] for l in leaves]
    del wb[SHEET_RECORD]
    ws = wb.create_sheet(SHEET_RECORD, 0)
    ws.cell(row=1, column=1, value="日期")
    for j, n in enumerate(names, start=2):
        ws.cell(row=1, column=j, value=n)
    for i, d in enumerate(sorted(records.keys()), start=2):
        ws.cell(row=i, column=1, value=d)
        for j, n in enumerate(names, start=2):
            v = records[d].get(n)
            if v not in (None, ""):
                ws.cell(row=i, column=j, value=v)
    return ws


def _toggle_item(ws, items, item_id):
    it = _find(items, item_id)
    if it is None or it["type"] != TYPE_CHECK or it.get("status") == STATUS_ARCHIVED:
        raise KeyError("勾选项目不存在")
    col = _name_to_col(ws).get(it["name"])
    if col is None:
        raise KeyError("数据列缺失")
    r = _ensure_row(ws, date.today().isoformat())
    cell = ws.cell(row=r, column=col)
    new_val = 0 if cell.value in (1, "1") else 1
    cell.value = new_val
    return new_val


def _set_count(ws, items, item_id, count):
    """多次打卡项：设置当日次数（自动钳制到 0~目标次数）。"""
    it = _find(items, item_id)
    if it is None or it["type"] != TYPE_MULTI or it.get("status") == STATUS_ARCHIVED:
        raise KeyError("多次打卡项不存在")
    try:
        c = int(float(count))
    except (TypeError, ValueError):
        raise ValueError("次数必须为数值")
    c = min(max(c, 0), it.get("target", 1))
    col = _name_to_col(ws).get(it["name"])
    if col is None:
        raise KeyError("数据列缺失")
    r = _ensure_row(ws, date.today().isoformat())
    ws.cell(row=r, column=col).value = c
    return c


def _set_value(ws, items, item_id, value):
    it = _find(items, item_id)
    if it is None or it["type"] != TYPE_NUMBER or it.get("status") == STATUS_ARCHIVED:
        raise KeyError("数值项不存在")
    try:
        v = float(value)
    except (TypeError, ValueError):
        raise ValueError("数值无效")
    if v != v or abs(v) >= MAX_AMOUNT:
        raise ValueError("数值超出范围")
    v = int(v) if v.is_integer() else round(v, 2)
    col = _name_to_col(ws).get(it["name"])
    if col is None:
        raise KeyError("数据列缺失")
    r = _ensure_row(ws, date.today().isoformat())
    ws.cell(row=r, column=col).value = v
    return v


def _parse_record_date(ds):
    try:
        d = datetime.strptime(str(ds).strip(), "%Y-%m-%d").date()
    except (TypeError, ValueError):
        raise ValueError("日期格式应为 YYYY-MM-DD")
    if d > date.today():
        raise ValueError("不能补录未来的日期")
    return d


def _set_record(ws, items, item_id, ds, value):
    """给指定日期（含过去日期）设置某个项目的值，用于历史补录。"""
    it = _find(items, item_id)
    if it is None or it.get("status") == STATUS_ARCHIVED:
        raise KeyError("项目不存在")
    d = _parse_record_date(ds)
    col = _name_to_col(ws).get(it["name"])
    if col is None:
        raise KeyError("数据列缺失")
    r = _ensure_row(ws, d.isoformat())
    if it["type"] == TYPE_CHECK:
        new = 1 if _num(value) != 0 else 0
    elif it["type"] == TYPE_MULTI:
        try:
            c = int(float(value))
        except (TypeError, ValueError):
            raise ValueError("次数必须为数值")
        new = min(max(c, 0), it.get("target", 1))  # 封顶目标次数
    else:
        try:
            v = float(value)
        except (TypeError, ValueError):
            raise ValueError("数值无效")
        if v != v or abs(v) >= MAX_AMOUNT:
            raise ValueError("数值超出范围")
        new = int(v) if v.is_integer() else round(v, 2)
    ws.cell(row=r, column=col).value = new
    _sort_rows(ws)
    return new


def _apply_day(ws, items, ds, changes):
    """批量补录某一天：changes = [{"id": 项目ID, "value": 值}, ...]"""
    d = _parse_record_date(ds)
    _ensure_row(ws, d.isoformat())
    result = {}
    for ch in changes:
        item_id = int(ch.get("id"))
        v = _set_record(ws, items, item_id, d.isoformat(), ch.get("value", 0))
        result[item_id] = v
    return result


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _leaf_done(leaf, v):
    """叶子项在给定原始值下是否算「当天完成」：勾选=1；多次项=次数 ≥ 目标次数。"""
    if leaf["type"] == TYPE_CHECK:
        return v in (1, "1")
    if leaf["type"] == TYPE_MULTI:
        t = leaf.get("target", 1) or 1
        return _num(v) >= t
    return False


def _leaf_gain(leaf, v):
    """叶子项当日贡献的收益：勾选/多次项达标才结算一次固定额；数值项=数值本身。"""
    if leaf["type"] in (TYPE_CHECK, TYPE_MULTI):
        return float(leaf.get("benefit") or 0) if _leaf_done(leaf, v) else 0.0
    return _num(v) if v not in (None, "") else 0.0


# ---------------- 收益（每日收益表） ----------------

def _read_benefits(wb):
    if SHEET_BENEFIT not in wb.sheetnames:
        return {}
    ws = wb[SHEET_BENEFIT]
    out = {}
    for r in range(2, ws.max_row + 1):
        d = ws.cell(row=r, column=1).value
        if d in (None, ""):
            continue
        out[str(d).strip()] = (round(_num(ws.cell(row=r, column=2).value), 2),
                               round(_num(ws.cell(row=r, column=3).value), 2))
    return out


def _daily_gain(records, items):
    """
    当日新增收益 = Σ(勾选项勾选×收益) + Σ(多次项达标×收益) + Σ(数值项当日数值)。
    """
    leaves = _leaf_order(items)
    gains = {}
    for d, vals in records.items():
        g = 0.0
        for leaf in leaves:
            g += _leaf_gain(leaf, vals.get(leaf["name"]))
        gains[d] = round(g, 2)
    return gains


def _sync_benefits(wb, items):
    """
    生成/校正每日收益表：累计收益 = 上一日累计 + 当日新增；
    首日累计 = 初始值(-7945) + 当日新增。返回是否有改动。
    数据源为主表 + 冷表合并视图（归档项的历史收益照常计入）。
    """
    records, _ = _read_all_records(wb, items)
    gains = _daily_gain(records, items)
    existing = _read_benefits(wb)
    dates = sorted(set(records.keys()) | set(existing.keys())) or [date.today().isoformat()]

    rows = []
    total = None
    for d in dates:
        g = gains.get(d)
        if g is None:
            g = existing.get(d, (0.0, 0.0))[0]
        total = INITIAL_BENEFIT if total is None else total
        total = round(total + g, 2)
        rows.append((d, round(g, 2), total))

    new_map = {d: (g, t) for d, g, t in rows}
    old_map = existing
    if new_map == old_map and SHEET_BENEFIT in wb.sheetnames:
        return False

    if SHEET_BENEFIT in wb.sheetnames:
        del wb[SHEET_BENEFIT]
    ws = wb.create_sheet(SHEET_BENEFIT)
    for i, h in enumerate(BENEFIT_HEADERS, start=1):
        ws.cell(row=1, column=i, value=h)
    for i, (d, g, t) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=d)
        ws.cell(row=i, column=2, value=g)
        ws.cell(row=i, column=3, value=t)
    return True


def _sync_all(wb, items):
    """同步打卡记录表 + 每日收益表；旧文件的归档列首次同步时搬入冷表。"""
    changed = False
    if _migrate_legacy_archived(wb, items):
        _rewrite_records(wb, items)   # 归档列已入冷表，主表重建为清一色在役列
        changed = True
    changed = _sync_records(_sheet(wb), items) or changed
    changed = _sync_benefits(wb, items) or changed
    return changed


def _compute_streak(records, leaves):
    """连续打卡天数：从今天（今天暂无完成则从昨天往回）起，连续至少完成一项的天数。"""
    cur = date.today()

    def done_on(d):
        vals = records.get(d.isoformat(), {})
        return any(_leaf_done(l, vals.get(l["name"])) for l in leaves)

    if not done_on(cur):
        cur -= timedelta(days=1)
    streak = 0
    while done_on(cur):
        streak += 1
        cur -= timedelta(days=1)
    return streak


def _build_state(wb, items):
    records, _ = _read_all_records(wb, items)
    today = date.today().isoformat()
    today_vals = records.get(today, {})
    benefits = _read_benefits(wb)
    leaves = _leaf_order(items)
    countable = [l for l in leaves
                 if l["type"] in (TYPE_CHECK, TYPE_MULTI)
                 and l.get("status") != STATUS_ARCHIVED]  # 归档项退出日常打卡
    tree = _build_tree(items)

    def attach(nodes):
        done = total = 0
        for n in nodes:
            if n["type"] == TYPE_GROUP:
                d, t = attach(n["children"])
                n["done"], n["total"] = d, t
                done += d
                total += t
            elif n["type"] == TYPE_CHECK:
                n["value"] = 1 if today_vals.get(n["name"]) in (1, "1") else 0
                total += 1
                done += n["value"]
            elif n["type"] == TYPE_MULTI:
                n["value"] = min(max(int(_num(today_vals.get(n["name"]))), 0), n.get("target", 1))
                total += 1
                done += 1 if n["value"] >= (n.get("target") or 1) else 0
            else:
                raw = today_vals.get(n["name"])
                n["value"] = _num(raw) if raw not in (None, "") else 0
        return done, total

    done, total = attach(tree)
    gain, cum = benefits.get(today, (0.0, INITIAL_BENEFIT))

    # 数据体检：记录表出现结构里没有的列，通常是手工改了 Excel 的结构表
    orphans = _orphan_columns(_sheet(wb), items)
    warning = None
    if orphans:
        names = "、".join(orphans[:5]) + ("等" if len(orphans) > 5 else "")
        warning = ("记录表存在 %d 个孤立列（%s），结构表中没有对应项目——"
                   "很可能直接改过 data 里的 Excel。请通过「结构管理」维护项目，"
                   "历史数据不会丢，但不同步的列将不参与统计" % (len(orphans), names))

    yesterday_d = (date.today() - timedelta(days=1)).isoformat()
    yvals = records.get(yesterday_d, {})
    ydone = sum(1 for l in countable if _leaf_done(l, yvals.get(l["name"])))
    ygain = benefits.get(yesterday_d, (0.0, INITIAL_BENEFIT))[0]
    return {
        "today": today,
        "weekday": WEEKDAYS_CN[date.today().weekday()],
        "tree": tree,
        "archived": _build_tree(items, archived=True),  # 归档箱（软删除项目）
        "done": done,
        "total": total,
        "warning": warning,                        # 数据体检提示（如孤立列）
        "streak": _compute_streak(records, countable),
        "benefit": {"gain": round(gain, 2), "total": round(cum, 2),
                    "initial": INITIAL_BENEFIT},
        "yesterday": {
            "date": yesterday_d,
            "exists": yesterday_d in records,
            "done": ydone,
            "total": len(countable),
            "gain": round(ygain, 2),
            "complete": len(countable) > 0 and ydone == len(countable),
        },
    }


# ---------------- 数据分析 ----------------

def _heatmap(records, benefits, leaves, days=HEATMAP_DAYS):
    """最近 days 天每天的完成情况（用于热力图）。leaves 只需传勾选项+多次项。"""
    today = date.today()
    out = []
    for i in range(days - 1, -1, -1):
        d = (today - timedelta(days=i)).isoformat()
        vals = records.get(d, {})
        done = sum(1 for l in leaves if _leaf_done(l, vals.get(l["name"])))
        gain = benefits.get(d, (0.0, 0.0))[0]
        out.append({"date": d, "done": done, "total": len(leaves),
                    "gain": round(gain, 2)})
    return out


def _trend(records, benefits, leaves, kind, count):
    """按周（周一起）/ 月聚合完成率与收益，返回最近 count 档（升序）。"""
    buckets = {}
    for d, vals in records.items():
        try:
            dd = datetime.strptime(d, "%Y-%m-%d").date()
        except (TypeError, ValueError):
            continue
        if kind == "week":
            key = dd - timedelta(days=dd.weekday())
            label = key.strftime("%m/%d")
        else:
            key = dd.replace(day=1)
            label = "%d月" % key.month
        b = buckets.setdefault(key, {"label": label, "done": 0, "total": 0, "gain": 0.0})
        b["total"] += len(leaves)
        b["done"] += sum(1 for l in leaves if _leaf_done(l, vals.get(l["name"])))
        b["gain"] = round(b["gain"] + benefits.get(d, (0.0, 0.0))[0], 2)

    today = date.today()
    keys = []
    if kind == "week":
        k = today - timedelta(days=today.weekday())
        for _ in range(count):
            keys.append(k)
            k -= timedelta(days=7)
    else:
        k = today.replace(day=1)
        for _ in range(count):
            keys.append(k)
            k = (k.replace(day=1) - timedelta(days=1)).replace(day=1)
    out = []
    for k in sorted(keys):
        b = buckets.get(k)
        if b is None:
            label = (k.strftime("%m/%d") if kind == "week" else "%d月" % k.month)
            out.append({"label": label, "done": 0, "total": 0, "gain": 0.0, "rate": 0})
        else:
            out.append({"label": b["label"], "done": b["done"], "total": b["total"],
                        "gain": b["gain"],
                        "rate": round(b["done"] / b["total"], 4) if b["total"] else 0})
    return out


def _month_of(ds):
    """日期字符串 -> 自然月 'YYYY-MM'。"""
    return str(ds)[:7]


def _month_elapsed_days(month, today_iso):
    """
    自然月已过天数（月度统计分母）：
      - 当月：今天日期号（如 10-17 -> 17）
      - 已过去的整月：该月自然天数（28/29/30/31）
    """
    if month == today_iso[:7]:
        try:
            return max(int(today_iso[8:10]), 1)
        except ValueError:
            return 1
    try:
        return monthrange(int(month[:4]), int(month[5:7]))[1]
    except ValueError:
        return 0


def _month_item_stat(leaf, path, month, vals, elapsed, prev=None):
    """
    单个项目在单个自然月内的统计。vals = [(日期, 原始值), ...]（该月内的记录）。
    month = "all" 表示全历史累计。字段与月度统计保持同构，便于前端复用渲染。
    """
    t = leaf.get("target", 1) or 1
    done_days = times = active_days = 0
    total_v = 0.0
    gain = 0.0
    for _d, v in vals:
        if leaf["type"] == TYPE_CHECK:
            if v in (1, "1"):
                done_days += 1
                gain += float(leaf.get("benefit") or 0)
        elif leaf["type"] == TYPE_MULTI:
            c = min(max(int(_num(v)), 0), t)
            times += c
            if c >= t:
                done_days += 1
                gain += float(leaf.get("benefit") or 0)
        else:
            x = _num(v) if v not in (None, "") else 0.0
            total_v = round(total_v + x, 2)
            gain += x
            if x > 0:
                active_days += 1
    s = {"id": leaf["id"], "name": leaf["name"], "path": path, "type": leaf["type"],
         "benefit": leaf.get("benefit", 0), "target": t,
         "archived": leaf.get("status") == STATUS_ARCHIVED,
         "month": month,
         "done_days": done_days, "total_days": elapsed,
         "rate": round(done_days / elapsed, 4) if elapsed else 0,
         "gain": round(gain, 2),
         "times": times,
         "avg_times": round(times / elapsed, 2) if elapsed else 0,
         "total": total_v if leaf["type"] == TYPE_NUMBER else 0,
         "avg": round(total_v / elapsed, 2)
                if (elapsed and leaf["type"] == TYPE_NUMBER) else 0,
         "days": active_days}
    if prev is not None:
        s["prev_done_days"] = prev["done_days"]
        s["prev_rate"] = prev["rate"]
        s["prev_gain"] = prev["gain"]
        s["prev_times"] = prev["times"]
        s["prev_total"] = prev["total"]
    return s


def _contrib(records, items):
    """分类贡献：按顶级分类聚合其下所有叶子项产生的收益与活跃天数。"""
    by_id = {it["id"]: it for it in items}

    def top(it):
        cur, guard = it, 0
        while cur["parent"] is not None and cur["parent"] in by_id and guard < 20:
            cur = by_id[cur["parent"]]
            guard += 1
        return cur

    gains, totals, actives, tops = {}, {}, {}, {}
    for d, vals in records.items():
        for l in _leaf_order(items):
            v = vals.get(l["name"])
            if l["type"] in (TYPE_CHECK, TYPE_MULTI):
                is_active = _leaf_done(l, v)
                g = float(l.get("benefit") or 0) if is_active else 0.0
            else:
                g = _num(v) if v not in (None, "") else 0.0
                is_active = g != 0
            t = top(l)
            name = t["name"]
            tops[name] = t
            gains[name] = round(gains.get(name, 0.0) + g, 2)
            totals[name] = totals.get(name, 0) + 1
            if is_active:
                actives[name] = actives.get(name, 0) + 1
    out = [{"name": k, "gain": round(gains[k], 2), "active_days": actives.get(k, 0),
            "total_days": totals[k],
            "archived": tops[k].get("status") == STATUS_ARCHIVED}
           for k in gains]
    out.sort(key=lambda x: x["gain"], reverse=True)
    return out


def _period_stat(records, benefits, countable, start, end, today_iso):
    """
    [start, end] 自然周期（周：周一~周日；月：1号~月末）内的打卡统计。
    只统计已过日期（<= today_iso）；未来日期在 days 中以 future 标记（供前端画空位）。
    返回每日明细、完成数、收益汇总、最佳项目。环比口径（日均）由调用方计算。
    """
    days = []
    done = gain = 0
    counts = {l["name"]: 0 for l in countable}
    d = start
    while d <= end:
        ds = d.isoformat()
        if ds > today_iso:                      # 周期内尚未到来的日期
            days.append({"date": ds, "future": True, "done": 0,
                         "total": len(countable), "gain": 0.0})
            d += timedelta(days=1)
            continue
        vals = records.get(ds, {})
        d_done = sum(1 for l in countable if _leaf_done(l, vals.get(l["name"])))
        for l in countable:
            v = vals.get(l["name"])
            if l["type"] == TYPE_MULTI:
                counts[l["name"]] += min(max(int(_num(v)), 0), l.get("target", 1))
            elif v in (1, "1"):
                counts[l["name"]] += 1
        days.append({"date": ds, "done": d_done, "total": len(countable),
                     "gain": round(benefits.get(ds, (0.0, 0.0))[0], 2)})
        done += d_done
        gain += benefits.get(ds, (0.0, 0.0))[0]
        d += timedelta(days=1)
    elapsed = sum(1 for x in days if not x.get("future"))
    total = elapsed * len(countable)
    best_name, best_cnt = "", 0
    for nm, c in counts.items():
        if c > best_cnt:
            best_name, best_cnt = nm, c
    return {"start": start.isoformat(), "end": end.isoformat(),
            "elapsed": elapsed, "days": days, "done": done, "total": total,
            "gain": round(gain, 2),
            "avg_gain": round(gain / elapsed, 2) if elapsed else 0.0,
            "best": {"name": best_name, "count": best_cnt}}


def _month_balance(benefits, today, month):
    """月末余额：上月末落定余额 -> 当前余额 -> 按已过日均预计月末余额。"""
    today_iso = today.isoformat()
    first = today.replace(day=1)
    prev_end = first - timedelta(days=1)
    prev_balance = benefits.get(prev_end.isoformat(), (0.0, INITIAL_BENEFIT))[1]
    current = benefits.get(today_iso, (0.0, INITIAL_BENEFIT))[1]
    elapsed = _month_elapsed_days(month, today_iso)
    m_gain = sum(g for d, (g, _) in benefits.items() if _month_of(d) == month)
    avg = m_gain / elapsed if elapsed else 0.0
    days_in_month = monthrange(today.year, today.month)[1]
    remaining = max(days_in_month - elapsed, 0)
    return {"prev_month_end": prev_end.isoformat(),
            "prev_balance": round(prev_balance, 2),
            "current": round(current, 2),
            "month_gain": round(m_gain, 2),
            "avg_daily": round(avg, 2),
            "remaining_days": remaining,
            "forecast": round(current + avg * remaining, 2)}


def _build_report(records, benefits, countable):
    """周报（本周一~周日）+ 月报（本月）+ 月末余额。环比统一用「日均收益」口径。"""
    today = date.today()
    today_iso = today.isoformat()
    # 周报：本周一 ~ 本周日；上周同期用于环比
    monday = today - timedelta(days=today.weekday())
    last_monday = monday - timedelta(days=7)
    week = _period_stat(records, benefits, countable, monday,
                        monday + timedelta(days=6), today_iso)
    last_week = _period_stat(records, benefits, countable, last_monday,
                             monday - timedelta(days=1), today_iso)
    # 月报：本月 1 号 ~ 月末；上月整月用于环比
    first = today.replace(day=1)
    month_end = today.replace(day=monthrange(today.year, today.month)[1])
    prev_first = (first - timedelta(days=1)).replace(day=1)
    prev_end = first - timedelta(days=1)
    month = _period_stat(records, benefits, countable, first, month_end, today_iso)
    last_month = _period_stat(records, benefits, countable, prev_first, prev_end, today_iso)
    return {
        "week": week,
        "week_prev_avg": last_week["avg_gain"],
        "month": month,
        "month_prev_avg": last_month["avg_gain"],
        "month_balance": _month_balance(benefits, today, _month_of(today_iso)),
        "streak": _compute_streak(records, countable),
    }


def _build_history(wb, items):
    # 主表 + 冷表合并：归档项的历史列、历史统计与收益链完整保留
    records, _ = _read_all_records(wb, items)
    leaves = _leaf_order(items)
    paths = _paths(items)
    columns = [{"id": l["id"], "name": l["name"], "type": l["type"],
                "benefit": l.get("benefit", 0), "target": l.get("target", 1),
                "path": paths[l["id"]],
                "archived": l.get("status") == STATUS_ARCHIVED} for l in leaves]
    benefits = _read_benefits(wb)
    countable = [l for l in leaves if l["type"] in (TYPE_CHECK, TYPE_MULTI)]
    today_iso = date.today().isoformat()
    sorted_dates = sorted(records.keys())
    rows = []
    for d in sorted(records.keys(), reverse=True):
        vals = records[d]
        row_vals = []
        for l in leaves:
            v = vals.get(l["name"])
            if l["type"] == TYPE_CHECK:
                row_vals.append(1 if v in (1, "1") else 0)
            elif l["type"] == TYPE_MULTI:
                row_vals.append(min(max(int(_num(v)), 0), l.get("target", 1)))
            else:
                row_vals.append(_num(v) if v not in (None, "") else 0)
        gain, cum = benefits.get(d, (0.0, 0.0))
        rows.append({"date": d, "values": row_vals, "gain": round(gain, 2),
                     "total": round(cum, 2)})
    total_days = len(rows)

    # 各项目全历史统计（month="all"，作为「全部历史」视图）
    all_vals = {l["id"]: [(d, records[d].get(l["name"])) for d in sorted_dates]
                for l in leaves}
    stats = [_month_item_stat(l, paths[l["id"]], "all", all_vals[l["id"]],
                              total_days) for l in leaves]

    # 各项目自然月统计（按月聚合，含环比）
    months = sorted({_month_of(d) for d in records} |
                    {_month_of(d) for d in benefits})
    monthly = {}
    month_meta = []
    for mi, month in enumerate(months):
        elapsed = _month_elapsed_days(month, today_iso)
        m_dates = [d for d in sorted_dates if _month_of(d) == month]
        prev_entries = monthly.get(months[mi - 1]) if mi > 0 else None
        entries = []
        for i, l in enumerate(leaves):
            vals = [(d, records[d].get(l["name"])) for d in m_dates]
            entries.append(_month_item_stat(
                l, paths[l["id"]], month, vals, elapsed,
                prev=prev_entries[i] if prev_entries else None))
        monthly[month] = entries
        gain = round(sum(benefits.get(d, (0.0, 0.0))[0] for d in m_dates), 2)
        month_end = benefits.get(m_dates[-1], (0.0, INITIAL_BENEFIT))[1] if m_dates \
            else INITIAL_BENEFIT
        month_meta.append({"month": month, "days": elapsed,
                           "record_days": len(m_dates), "gain": gain,
                           "total": round(month_end, 2)})

    benefit_rows = [{"date": d, "gain": round(g, 2), "total": round(t, 2)}
                    for d, (g, t) in sorted(benefits.items())]
    return {
        "columns": columns,
        "rows": rows,
        "stats": stats,
        "months": months,          # 升序，如 ["2026-08", "2026-09", "2026-10"]
        "month_meta": month_meta,  # 每月一行汇总（含月末累计收益）
        "monthly": monthly,        # month -> [项目统计...]
        "benefits": benefit_rows,
        "heatmap": _heatmap(records, benefits, countable),
        "weekly": _trend(records, benefits, countable, "week", TREND_WEEKS),
        "monthly_trend": _trend(records, benefits, countable, "month", TREND_MONTHS),
        "report": _build_report(records, benefits, countable),
        "contrib": _contrib(records, items),
    }


# ---------------- 路由 ----------------

def _save_error():
    return jsonify({"ok": False, "error": "保存失败：请先关闭 data/打卡记录.xlsx 后重试"}), 500


@app.route("/")
def index():
    return render_template("index.html")


def _load_synced():
    """加载工作簿并同步（保存失败时返回 (wb, items, warning)）。"""
    wb = _load()
    items = _read_structure(wb)
    warning = None
    if _sync_all(wb, items):
        try:
            _save(wb)
        except PermissionError as e:
            warning = str(e)
    return wb, items, warning


@app.route("/api/state")
def api_state():
    with _lock:
        wb, items, warning = _load_synced()
        payload = _build_state(wb, items)
    if warning:
        payload["warning"] = warning
    return jsonify(payload)


@app.route("/api/history")
def api_history():
    with _lock:
        wb, items, warning = _load_synced()
        payload = _build_history(wb, items)
    if warning:
        payload["warning"] = warning
    return jsonify(payload)


@app.route("/api/toggle", methods=["POST"])
def api_toggle():
    try:
        item_id = int((request.get_json(silent=True) or {}).get("id"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "缺少项目 ID"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            _sync_all(wb, items)
            value = _toggle_item(_sheet(wb), items, item_id)
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "value": value, "state": _build_state(wb, items)})
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/value", methods=["POST"])
def api_value():
    data = request.get_json(silent=True) or {}
    try:
        item_id = int(data.get("id"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "缺少项目 ID"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            _sync_all(wb, items)
            value = _set_value(_sheet(wb), items, item_id, data.get("value"))
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "value": value, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/count", methods=["POST"])
def api_count():
    """多次打卡项：设置当日次数（点第 k 个圆点 = 记 k 次，自动钳制 0~目标次数）。"""
    data = request.get_json(silent=True) or {}
    try:
        item_id = int(data.get("id"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "缺少项目 ID"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            _sync_all(wb, items)
            value = _set_count(_sheet(wb), items, item_id, data.get("count", 0))
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "value": value, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/day", methods=["POST"])
def api_day():
    """历史补录：批量设置某一天各项目的值，收益链自动重算。"""
    data = request.get_json(silent=True) or {}
    ds = data.get("date")
    changes = data.get("changes")
    if not isinstance(changes, list) or not changes:
        return jsonify({"ok": False, "error": "没有需要保存的修改"}), 400
    if len(changes) > MAX_CHANGES:
        return jsonify({"ok": False, "error": "一次最多修改 %d 项" % MAX_CHANGES}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            _sync_all(wb, items)
            _apply_day(_sheet(wb), items, ds, changes)
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


def _parent_from(data):
    parent_id = data.get("parent_id")
    try:
        return int(parent_id) if parent_id not in (None, "") else None
    except (TypeError, ValueError):
        raise ValueError("父级 ID 无效")


@app.route("/api/item/add", methods=["POST"])
def api_item_add():
    data = request.get_json(silent=True) or {}
    name = str(data.get("name", "")).strip()
    typ = str(data.get("type", "")).strip()
    benefit = data.get("benefit", 0)
    target = data.get("target", 1)
    try:
        parent_id = _parent_from(data)
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            _add_item(items, name, typ, parent_id, benefit, target)
            _write_structure(wb, items)
            ws = _rewrite_records(wb, items)
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "父级不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/item/update", methods=["POST"])
def api_item_update():
    data = request.get_json(silent=True) or {}
    name = str(data.get("name", "")).strip()
    benefit = data.get("benefit")
    target = data.get("target")
    type_ = data.get("type")
    if type_ is not None:
        type_ = str(type_).strip() or None
    try:
        item_id = int(data.get("id"))
        parent_id = _parent_from(data)
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "参数无效"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            old_items = [dict(it) for it in items]
            old_type = _update_item(items, item_id, name, parent_id, benefit, target, type_)
            _write_structure(wb, items)
            ws = _rewrite_records(wb, items, _rename_map(old_items, items))
            converted = bool(type_ and type_ != old_type)
            if converted:
                _convert_type_records(ws, items, item_id, old_type, type_)
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "converted": converted,
                            "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/item/rename", methods=["POST"])
def api_item_rename():
    """重命名项目（含归档项）：同步改写记录表列名，历史数据完整保留。"""
    data = request.get_json(silent=True) or {}
    try:
        item_id = int(data.get("id"))
        name = str(data.get("name", "")).strip()
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "参数无效"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            _migrate_legacy_archived(wb, items)   # 按旧名先把归档列搬入冷表
            old_items = [dict(it) for it in items]
            _rename_item(items, item_id, name)
            _write_structure(wb, items)
            _sync_cold_name(wb, item_id, name)    # 冷表名称快照同步
            _rewrite_records(wb, items, _rename_map(old_items, items))
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/item/remove", methods=["POST"])
def api_item_remove():
    """软删除：归档整棵子树。历史数据整列搬入冷表，收益链完全不动。"""
    try:
        item_id = int((request.get_json(silent=True) or {}).get("id"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "缺少项目 ID"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            ids = _archive_subtree(items, item_id)
            _write_structure(wb, items)
            _migrate_to_cold(wb, items, ids)   # 历史列 -> 冷表
            _rewrite_records(wb, items)        # 主表移除归档列，只剩在役项
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/item/restore", methods=["POST"])
def api_item_restore():
    """从归档箱恢复：连同祖先（向上闭合）与后代一起恢复，冷表数据搬回主表。"""
    try:
        item_id = int((request.get_json(silent=True) or {}).get("id"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "缺少项目 ID"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            ids = _restore_subtree(items, item_id)
            _write_structure(wb, items)
            _rewrite_records(wb, items)         # 补齐恢复项的列（空列）
            _restore_from_cold(wb, items, ids)  # 冷表历史值写回，删冷表行
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/item/purge", methods=["POST"])
def api_item_purge():
    """永久删除：从结构表移除，删除主表对应列与冷表历史行，并重算收益链。"""
    try:
        item_id = int((request.get_json(silent=True) or {}).get("id"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "缺少项目 ID"}), 400
    try:
        with _lock:
            wb = _load()
            items = _read_structure(wb)
            items, ids = _purge_subtree(items, item_id)
            _write_structure(wb, items)
            _rewrite_records(wb, items)   # 列被物理移除
            _purge_cold(wb, ids)          # 冷表历史行同步移除
            _sync_all(wb, items)
            _save(wb)
            return jsonify({"ok": True, "state": _build_state(wb, items)})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except KeyError:
        return jsonify({"ok": False, "error": "项目不存在"}), 404
    except PermissionError:
        return _save_error()


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings():
    if request.method == "GET":
        return jsonify({"ok": True, "settings": _load_settings()})
    data = request.get_json(silent=True) or {}
    try:
        goal = data.get("goal", 0)
        goal = 0.0 if goal in (None, "") else float(goal)
        if goal != goal or abs(goal) >= MAX_AMOUNT:
            raise ValueError("目标数值超出范围")
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "目标数值无效"}), 400
    s = {"goal": round(goal, 2), "note": str(data.get("note", "") or "")[:30]}
    try:
        _save_settings(s)
    except OSError:
        return jsonify({"ok": False, "error": "设置保存失败"}), 500
    return jsonify({"ok": True, "settings": s})


@app.route("/api/backups")
def api_backups():
    return jsonify({"ok": True, "backups": _list_backups()})


@app.route("/api/backup", methods=["POST"])
def api_backup():
    try:
        dst = _backup(force=True)
        if not dst:
            return jsonify({"ok": False, "error": "备份失败：数据文件不存在"}), 400
        return jsonify({"ok": True, "name": os.path.basename(dst), "backups": _list_backups()})
    except OSError:
        return jsonify({"ok": False, "error": "备份失败"}), 500


def _export_zip_bytes():
    """把当前数据文件 + 全部历史备份打包成 zip（内存中生成，返回 bytes）。

    用途：一键导出备份包，方便把数据带出这台电脑（U 盘 / 网盘 / 文件传输助手）。
    系统内备份都在同一块硬盘上，救不了硬盘损坏或整个项目被误删，导出才是有副本。
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        files = []
        if os.path.exists(XLSX_PATH):
            files.append((XLSX_PATH, "打卡记录.xlsx"))
        for f in sorted(glob.glob(os.path.join(_backup_dir(), "打卡记录_*.xlsx"))):
            files.append((f, os.path.join("backup", os.path.basename(f))))
        if not files:
            raise OSError("没有可导出的数据文件")
        notes = (
            "打卡系统备份包\n"
            "导出时间：%s\n\n"
            "内容：\n"
            "  打卡记录.xlsx   当前正在使用的数据（最重要，恢复用这一份）\n"
            "  backup/         系统自动保存的历史备份（每天一份，最多 30 份）\n\n"
            "如何恢复：\n"
            "  1. 先关闭打卡系统（黑色命令行窗口）\n"
            "  2. 解压本包，用「打卡记录.xlsx」覆盖 data 文件夹里的同名文件\n"
            "  3. 重新双击 start.bat 启动即可\n\n"
            "建议：每周导出一次，保存到 U 盘、移动硬盘或网盘上。\n"
        ) % datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        zf.writestr("先读我_恢复说明.txt", notes)
        for src, arc in files:
            zf.write(src, arc)
    return buf.getvalue()


@app.route("/api/export/zip")
def api_export_zip():
    try:
        data = _export_zip_bytes()
    except OSError:
        return jsonify({"ok": False, "error": "导出失败：没有找到数据文件"}), 400
    name = "打卡备份_%s.zip" % datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(io.BytesIO(data), mimetype="application/zip",
                     as_attachment=True, download_name=name)


if __name__ == "__main__":
    with _lock:
        wb = _load()
        items = _read_structure(wb)
        if _sync_all(wb, items):
            _save(wb)
    port = int(os.environ.get("PORT", 5000))
    print("=" * 48)
    print("  每日打卡记录系统已启动")
    print("  浏览器访问: http://127.0.0.1:%d" % port)
    print("  数据文件: %s" % XLSX_PATH)
    print("  按 Ctrl+C 停止服务")
    print("=" * 48)
    app.run(host="127.0.0.1", port=port, debug=False)
