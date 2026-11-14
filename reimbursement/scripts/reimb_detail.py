#!/usr/bin/env python3
"""统一「报销明细」sheet 写入（2026-11-14 23:53 用户新口径）。

用户原话：「为了兼容表格数据，你可以在表格内新建一个报销明细。这样就不会有兼容的问题了。
所有的表单统一在表内建一个报销明细。」

为什么要 zip 层精确插入，而不是 openpyxl 整体 save：
  司机工资表是 WPS 生成的 xlsx，含 customXml/、tables/、comments、vmlDrawing、
  dbsheet 命名空间等 openpyxl 不认识或不保真的部件。整体重写会丢特性、改公式引用。
  本脚本只改 4 个清单类部件并新增 1 个 worksheet 部件，其余所有部件字节原样复制。

写入目标统一为「报销明细」sheet，不再去猜各司机原表里位置不一的「报销 / 报销支出」
汇总单元格（歧义 / 公式 / 无数值等兼容问题由此彻底消失）。

用法：
  python3 scripts/reimb_detail.py apply --data <data.json>
  python3 scripts/reimb_detail.py inspect --file <xlsx>

data.json:
{
  "run_id": "2026-11-14_2353",
  "year": 2026,
  "backup_dir": "state/backups/2026-11-14_2353",
  "precreate_all": true,          // 给该年份目录下所有司机报表预建「报销明细」
  "entries": [                    // 本轮新增入账条目，按 driver 分组追加
    {"driver":"王博成","key":"om_xxx","receipt_date":"2026-11-11",
     "pay_time":"2026-11-11 12:49:53","amount":45.14,"order_no":"...",
     "category":"充电","merchant":"...","image_file":"...","message_time":"...",
     "entry_id":"30","task_url":"...","note":""}
  ]
}

输出 JSON：{"ok", "backup_dir", "files":[{file, action, rows_added, total_rows,
            total_amount, verify}], "errors":[]}
"""
import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from xml.sax.saxutils import escape as xml_escape

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "..", "config.json")
SHEET_NAME = "报销明细"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
HEADERS = ["唯一键", "司机", "消费日期", "支付时间", "金额(元)", "订单号", "费用类别",
           "商户/说明", "凭证文件名", "消息时间", "台账ID", "确认任务链接", "备注"]
# 金额列索引（0-based）：E 列 = 4
AMOUNT_COL = 4
KEY_COL = 0


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_str(ref, value):
    return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        ref, xml_escape("" if value is None else str(value)))


def cell_num(ref, value):
    try:
        v = float(value)
    except (TypeError, ValueError):
        return cell_str(ref, value)
    # 避免科学计数法与浮点尾巴
    s = ("%.2f" % v).rstrip("0").rstrip(".") if v != int(v) else str(int(v))
    return '<c r="%s"><v>%s</v></c>' % (ref, s or "0")


def row_xml(r, values, amount_col=AMOUNT_COL):
    cells = []
    for i, v in enumerate(values):
        ref = "%s%d" % (col_letter(i), r)
        if i == amount_col and isinstance(v, (int, float)):
            cells.append(cell_num(ref, v))
        else:
            cells.append(cell_str(ref, v))
    return '<row r="%d">%s</row>' % (r, "".join(cells))


def entry_values(e, driver):
    return [
        e.get("key") or e.get("message_id") or "",
        driver,
        e.get("receipt_date") or "",
        e.get("pay_time") or "",
        e.get("amount") or 0,
        e.get("order_no") or "",
        e.get("category") or "",
        e.get("merchant") or "",
        e.get("image_file") or "",
        e.get("message_time") or "",
        e.get("entry_id") or "",
        e.get("task_url") or "",
        e.get("note") or "",
    ]


def sheet_xml(rows):
    """rows: list of (row_number, values, is_total)"""
    widths = [30, 10, 12, 20, 12, 30, 14, 28, 46, 20, 9, 52, 30]
    cols = "".join('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i + 1, i + 1, w)
                   for i, w in enumerate(widths))
    body = "".join(row_xml(r, v) for r, v, _ in rows)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s"><dimension ref="A1:%s%d"/>'
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            '</sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            '<cols>%s</cols><sheetData>%s</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            '</worksheet>') % (NS_MAIN, col_letter(len(HEADERS) - 1), max(1, len(rows)), cols, body)


def total_row(r, amount):
    vals = ["合计", "", "", "", round(float(amount), 2)] + [""] * (len(HEADERS) - 5)
    return (r, vals, True)


def read_parts(path):
    with zipfile.ZipFile(path) as z:
        return [(i, z.read(i.filename)) for i in z.infolist()]


def find_detail_sheet(parts):
    """返回 (sheet_xml_part_name, rel_id) 或 (None, None)"""
    wb = next((d.decode("utf-8") for i, d in parts if i.filename == "xl/workbook.xml"), "")
    m = re.search(r'<sheet[^>]*name="%s"[^>]*/>' % re.escape(SHEET_NAME), wb)
    if not m:
        return None, None
    rid = re.search(r'r:id="([^"]+)"', m.group(0))
    rels = next((d.decode("utf-8") for i, d in parts
                 if i.filename == "xl/_rels/workbook.xml.rels"), "")
    if not rid:
        return None, None
    r = re.search(r'<Relationship[^>]*Id="%s"[^>]*Target="([^"]+)"' % re.escape(rid.group(1)), rels)
    if not r:
        return None, None
    target = r.group(1)
    part = "xl/" + target.lstrip("/") if not target.startswith("/") else target.lstrip("/")
    return part, rid.group(1)


def extract_existing_rows(xml_text):
    """从已有明细 sheet 提取 (row_xml, 唯一键, 金额)，用于幂等追加与合计重算。"""
    m = re.search(r"<sheetData>(.*?)</sheetData>", xml_text, re.S)
    if not m:
        return []
    out = []
    for rm in re.finditer(r"<row\b[^>]*>.*?</row>", m.group(1), re.S):
        raw = rm.group(0)
        first = re.search(r'<c r="A\d+"[^>]*>.*?<t[^>]*>(.*?)</t>', raw, re.S)
        key = first.group(1) if first else ""
        amt_m = re.search(r'<c r="E\d+"[^>]*>\s*<v>([^<]+)</v>', raw)
        amt = 0.0
        if amt_m:
            try:
                amt = float(amt_m.group(1))
            except ValueError:
                amt = 0.0
        out.append((raw, key, amt))
    return out


def next_ids(parts):
    wb = next((d.decode("utf-8") for i, d in parts if i.filename == "xl/workbook.xml"), "")
    rels = next((d.decode("utf-8") for i, d in parts
                 if i.filename == "xl/_rels/workbook.xml.rels"), "")
    sheet_ids = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)]
    rid_nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    existing_parts = {i.filename for i, _ in parts}
    k = max([int(x) for x in re.findall(r"xl/worksheets/sheet(\d+)\.xml", " ".join(existing_parts))] or [0]) + 1
    while "xl/worksheets/sheet%d.xml" % k in existing_parts:
        k += 1
    return max(sheet_ids or [0]) + 1, "rId%d" % (max(rid_nums or [0]) + 1), \
        "xl/worksheets/sheet%d.xml" % k


def patch_workbook(wb, sheet_id, rid):
    tag = '<sheet name="%s" sheetId="%d" r:id="%s"/>' % (SHEET_NAME, sheet_id, rid)
    if "</sheets>" not in wb:
        raise ValueError("workbook.xml 缺少 </sheets>")
    return wb.replace("</sheets>", tag + "</sheets>", 1)


def patch_rels(rels, rid, part):
    target = part[len("xl/"):] if part.startswith("xl/") else part
    tag = ('<Relationship Id="%s" Type="%s/worksheet" Target="%s"/>' % (rid, REL_NS, target))
    if "</Relationships>" not in rels:
        raise ValueError("workbook.xml.rels 缺少 </Relationships>")
    return rels.replace("</Relationships>", tag + "</Relationships>", 1)


def patch_content_types(ct, part):
    tag = ('<Override PartName="/%s" ContentType="application/vnd.openxmlformats-'
           'officedocument.spreadsheetml.worksheet+xml"/>' % part)
    if "</Types>" not in ct:
        raise ValueError("[Content_Types].xml 缺少 </Types>")
    return ct.replace("</Types>", tag + "</Types>", 1)


def patch_app_xml(app):
    """同步工作表数量与名称，避免 Excel/WPS 打开时提示修复。"""
    if not app:
        return app
    try:
        m = re.search(r'(<vt:vector size=")(\d+)(" baseType="lpstr">)', app)
        if not m:
            return app
        app = app[:m.start(2)] + str(int(m.group(2)) + 1) + app[m.end(2):]
        app = re.sub(r'(<vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)',
                     lambda x: x.group(1) + str(int(x.group(2)) + 1), app, count=1)
        app = app.replace("</vt:vector></TitlesOfParts>",
                          "<vt:lpstr>%s</vt:lpstr></vt:vector></TitlesOfParts>" % SHEET_NAME, 1)
    except Exception:
        pass
    return app


def write_zip(path, parts, replacements, additions):
    tmp = path + ".tmp_reimb"
    with zipfile.ZipFile(tmp, "w") as zout:
        for info, data in parts:
            zout.writestr(info, replacements.get(info.filename, data))
        for name, data in additions.items():
            zout.writestr(name, data)
    os.replace(tmp, path)


def verify(path, before_parts, changed_names, expect_rows, expect_amount):
    """回读校验：未改部件字节完全一致 + openpyxl 能打开 + 明细行数/金额一致。"""
    after = {i.filename: d for i, d in read_parts(path)}
    before = {i.filename: d for i, d in before_parts}
    diffs = [n for n in before if n not in changed_names and before.get(n) != after.get(n)]
    result = {"unchanged_parts": len(before) - len(changed_names),
              "unexpected_diffs": diffs, "openpyxl_ok": False,
              "detail_rows": 0, "detail_amount": 0.0, "sheet_present": False}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
        result["sheet_present"] = SHEET_NAME in wb.sheetnames
        ws = wb[SHEET_NAME]
        vals = list(ws.values)
        body = [r for r in vals[1:] if r and str(r[0] or "").strip() not in ("", "合计")]
        result["detail_rows"] = len(body)
        result["detail_amount"] = round(sum(float(r[AMOUNT_COL] or 0) for r in body), 2)
        result["openpyxl_ok"] = True
        result["sheet_count"] = len(wb.sheetnames)
        wb.close()
    except Exception as e:
        result["openpyxl_error"] = str(e)
    result["rows_match"] = result["detail_rows"] == expect_rows
    result["amount_match"] = abs(result["detail_amount"] - float(expect_amount)) < 0.005
    result["ok"] = (not diffs and result["openpyxl_ok"] and result["sheet_present"]
                    and result["rows_match"] and result["amount_match"])
    return result


def apply_one(path, entries, backup_dir):
    """在单个 xlsx 内建/写「报销明细」。返回 (result_dict, error_or_None)。"""
    parts = read_parts(path)
    detail_part, _ = find_detail_sheet(parts)
    created = detail_part is None
    before = {i.filename: d for i, d in parts}
    replacements = {}
    additions = {}
    changed = set()

    if created:
        sheet_id, rid, detail_part = next_ids(parts)
        for name, fn in (("xl/workbook.xml", lambda s: patch_workbook(s, sheet_id, rid)),
                         ("xl/_rels/workbook.xml.rels", lambda s: patch_rels(s, rid, detail_part)),
                         ("[Content_Types].xml", lambda s: patch_content_types(s, detail_part)),
                         ("docProps/app.xml", patch_app_xml)):
            if name in before:
                replacements[name] = fn(before[name].decode("utf-8")).encode("utf-8")
                changed.add(name)
        existing = []          # (row_xml, key, amount)
        additions[detail_part] = ""
        changed.add(detail_part)
    else:
        existing = extract_existing_rows(before[detail_part].decode("utf-8"))
        changed.add(detail_part)

    # 丢弃旧合计行；已有业务键做幂等
    existing = [x for x in existing if x[1] not in ("", "唯一键", "合计")]
    seen = {x[1] for x in existing}
    fresh = []
    for e in entries:
        k = e.get("key") or e.get("message_id") or ""
        if k in seen:
            continue
        seen.add(k)
        fresh.append(e)

    rows = []
    r = 1
    rows.append((r, HEADERS, False))
    r += 1
    amount_sum = 0.0
    for raw, key, amt in existing:
        # 保留原行 XML，仅把行号重写（行号从 2 连续递增）
        raw2 = re.sub(r'(<row\b[^>]*\br=")\d+(")', lambda m: m.group(1) + str(r) + m.group(2), raw, count=1)
        raw2 = re.sub(r'(<c r=")([A-Z]+)(\d+)(")',
                      lambda m: m.group(1) + m.group(2) + str(r) + m.group(4), raw2)
        rows.append((r, raw2, True))   # 标记 raw 模式
        amount_sum += amt
        r += 1
    for e in fresh:
        rows.append((r, entry_values(e, e.get("driver") or ""), False))
        amount_sum += float(e.get("amount") or 0)
        r += 1
    rows.append((r, total_row(r, amount_sum)[1], False))

    # 生成 sheetData：raw 模式的行直接插入
    body = []
    for rr, v, is_raw in rows:
        if is_raw:
            body.append(v)
        else:
            body.append(row_xml(rr, v))
    sheet = sheet_xml([])  # 仅取外壳
    sheet = re.sub(r"<sheetData>.*?</sheetData>", "<sheetData>%s</sheetData>" % "".join(body),
                   sheet, flags=re.S)
    sheet = re.sub(r'<dimension ref="A1:[A-Z]+\d+"',
                   '<dimension ref="A1:%s%d"' % (col_letter(len(HEADERS) - 1), len(rows)), sheet)
    if created:
        additions[detail_part] = sheet.encode("utf-8")
    else:
        replacements[detail_part] = sheet.encode("utf-8")

    # 备份（必须先成功）
    os.makedirs(backup_dir, exist_ok=True)
    backup = os.path.join(backup_dir, os.path.basename(path))
    shutil.copy2(path, backup)
    write_zip(path, parts, replacements, additions)

    expect_rows = len(existing) + len(fresh)
    v = verify(path, parts, changed, expect_rows, round(amount_sum, 2))
    if not v["ok"]:
        # 校验失败立即回滚
        shutil.copy2(backup, path)
        v["rolled_back"] = True
    return {
        "file": path, "action": "created" if created else "appended",
        "detail_part": detail_part, "rows_added": len(fresh),
        "rows_skipped_duplicate": len(entries) - len(fresh),
        "total_rows": expect_rows, "total_amount": round(amount_sum, 2),
        "backup": backup, "verify": v,
    }, (None if v["ok"] else "回读校验失败，已回滚：%s" % json.dumps(v, ensure_ascii=False)[:400])


def cmd_apply(args):
    cfg = load_config()
    with open(args.data, "r", encoding="utf-8") as f:
        d = json.load(f)
    run_id = d["run_id"]
    year = int(d.get("year") or 2026)
    root = os.path.join(cfg["report_root"], "%d年" % year, "司机工资")
    backup_dir = d.get("backup_dir") or os.path.join(cfg["state_dir"], "backups", run_id)
    if not os.path.isabs(backup_dir):
        backup_dir = os.path.join(os.path.dirname(HERE), backup_dir)
    entries = d.get("entries") or []
    by_driver = {}
    for e in entries:
        by_driver.setdefault(e["driver"], []).append(e)

    files = []
    errors = []
    if d.get("precreate_all", True):
        import glob
        paths = sorted(glob.glob(os.path.join(root, "*.xlsx")))
    else:
        paths = [os.path.join(root, "%s%d.xlsx" % (drv, year)) for drv in by_driver]
    for p in paths:
        if not os.path.isfile(p):
            errors.append({"file": p, "error": "文件不存在，跳过（绝不新建报表文件）"})
            continue
        drv = os.path.basename(p).replace("%d.xlsx" % year, "")
        res, err = apply_one(p, by_driver.get(drv, []), backup_dir)
        files.append(res)
        if err:
            errors.append({"file": p, "error": err})
    return {"ok": not errors, "run_id": run_id, "sheet": SHEET_NAME,
            "backup_dir": backup_dir, "files": files, "errors": errors}


def cmd_inspect(args):
    parts = read_parts(args.file)
    detail_part, rid = find_detail_sheet(parts)
    out = {"file": args.file, "sheet_exists": bool(detail_part),
           "detail_part": detail_part, "rel_id": rid}
    if detail_part:
        raw = next((d.decode("utf-8") for i, d in parts if i.filename == detail_part), "")
        rows = extract_existing_rows(raw)
        body = [x for x in rows if x[1] not in ("", "唯一键", "合计")]
        out["rows"] = len(body)
        out["amount"] = round(sum(x[2] for x in body), 2)
        out["sheets"] = re.findall(r'<sheet[^>]*name="([^"]+)"',
                                   next(d.decode("utf-8") for i, d in parts
                                        if i.filename == "xl/workbook.xml"))
    return out


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("apply")
    a.add_argument("--data", required=True)
    i = sub.add_parser("inspect")
    i.add_argument("--file", required=True)
    args = ap.parse_args()
    out = cmd_apply(args) if args.cmd == "apply" else cmd_inspect(args)
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0 if out.get("ok", True) else 1


if __name__ == "__main__":
    sys.exit(main())
