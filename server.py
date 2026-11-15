from __future__ import annotations

import json
import socket
import threading
import webbrowser
from calendar import monthrange
from datetime import date, datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries


BASE_DIR = Path(__file__).resolve().parent
SOURCE_LABEL = "MONITORING SLA MSFO SGN"
KNOWN_SOURCES = (
    Path(r"D:\OneDrive\1. PT. SAPUTRA GLOBAL NUSANTARA\5. REPORT - EXECUTIVE\2. MSFO\1.MONITORING SLA MSFO SGN.xlsx"),
    Path(r"D:\OneDrive\1. PT. SAPUTRA GLOBAL NUSANTARA\5. REPORT - EXECUTIVE\7. MONITORING SLA MSFO SGN.xlsx"),
    Path(r"D:\OneDrive\Documents\7. MONITORING SLA MSFO SGN.xlsx"),
)
SEARCH_ROOTS = (
    Path(r"D:\OneDrive\1. PT. SAPUTRA GLOBAL NUSANTARA\5. REPORT - EXECUTIVE"),
    Path(r"D:\OneDrive\Documents"),
    Path.home() / "OneDrive",
)
HOST = "127.0.0.1"
START_PORT = 8765
_current_source: Path | None = None
_cache: dict[str, object] = {"source": None, "mtime_ns": None, "payload": None}
_cache_lock = threading.Lock()
MONTH_NAMES = (
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
)


def json_value(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def number(value):
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def choose_newest_source(candidates) -> Path | None:
    files = {path.resolve() for path in candidates if path.is_file() and not path.name.startswith("~$")}
    preferred = [path for path in files if "uji" not in path.stem.lower()] or list(files)
    return max(preferred, key=lambda path: path.stat().st_mtime_ns, default=None)


def build_daily_report(workbook, source: Path, area_names: list[str]) -> dict[str, object]:
    normalize = lambda value: str(value).upper().replace(" ", "")
    counts = {}
    open_list = []
    with ZipFile(source) as archive:
        table_limits = {}
        for name in archive.namelist():
            if not name.startswith("xl/tables/table") or not name.endswith(".xml"):
                continue
            table = ElementTree.fromstring(archive.read(name))
            table_limits[normalize(table.attrib["displayName"]).replace("_", "")] = range_boundaries(table.attrib["ref"])[3]

    def bucket(month, area):
        year, month_number = map(int, month.split("-"))
        days = monthrange(year, month_number)[1]
        return counts.setdefault(month, {}).setdefault(area, {
            "daily_open": [0] * days,
            "daily_close": [0] * days,
            "total": 0,
            "open_tickets": 0,
            "close_tickets": 0,
            "tickets_mp": 0,
            "tickets_low_loss": 0,
            "tickets_change_add": 0,
            "low_loss_within": 0,
            "low_loss_over": 0,
            "change_add_within": 0,
            "change_add_over": 0,
            "mp_within": 0,
            "mp_over": 0,
            "within_sla": 0,
            "over_sla": 0,
            "mttr_seconds": 0,
            "mttr_count": 0,
        })

    def add_event(area, event_value, kind):
        if not isinstance(event_value, (date, datetime)):
            return
        month = event_value.strftime("%Y-%m")
        bucket(month, area)[kind][event_value.day - 1] += 1

    def as_datetime(value):
        if isinstance(value, datetime):
            return value
        if isinstance(value, date):
            return datetime(value.year, value.month, value.day)
        return None

    for marker, table_marker, header_row, close_index in (("FAT LOSS-", "FATLOSS", 2, 3), ("MP-", "MP", 1, 2)):
        sheets = {
            normalize(name.split("-", 1)[1]): name
            for name in workbook.sheetnames
            if marker in name.upper()
        }
        selected = {}
        order = 0
        for area in area_names:
            sheet_name = sheets.get(normalize(area))
            if not sheet_name:
                continue
            table_end = next(
                (end for name, end in table_limits.items() if table_marker in name and name.endswith(normalize(area))),
                workbook[sheet_name].max_row,
            )
            for row in workbook[sheet_name].iter_rows(
                min_row=header_row + 1, max_row=table_end, min_col=2, max_col=44, values_only=True
            ):
                ticket_id, opened, closed, status = row[11], row[1], row[close_index], row[4]
                if not ticket_id:
                    continue
                order += 1
                is_closed = str(status).strip().upper() in {"DONE", "CLOSE"}
                event = closed if is_closed else opened
                rank_date = as_datetime(event) or datetime.min
                is_fat = table_marker == "FATLOSS"
                label = str(row[4] if row[4] not in (None, "") else row[12 if is_fat else 13] or "").strip()
                info = {
                    "id": str(ticket_id).strip(), "area": area, "status": label,
                    "type": str(row[13]).strip() if is_fat else "MP",
                    "detail": str(row[22] if is_fat else f"{row[12] or ''} · {row[25] or ''}").strip(" ·"),
                    "opened": json_value(opened), "sla": str(row[7] or "").strip(),
                } if opened and label.upper() not in {"DONE", "CLOSE", ""} else None
                candidate = (
                    (is_closed, rank_date, order), area, opened, closed, is_closed,
                    str(row[7]).strip().upper(), row[13] if table_marker == "FATLOSS" else None,
                    row[42] if table_marker == "MP" else None, info,
                )
                key = str(ticket_id).strip().upper()
                if key not in selected or candidate[0] > selected[key][0]:
                    selected[key] = candidate
        for _, area, opened, closed, is_closed, sla, category, nro, info in selected.values():
            if info and not is_closed:
                open_list.append(info)
            add_event(area, opened, "daily_open")
            add_event(area, closed, "daily_close")
            period_event = as_datetime(closed if is_closed else opened) or as_datetime(opened)
            if not period_event:
                continue
            if table_marker == "MP":
                ticket_type = "mp"
            else:
                category_key = normalize(category)
                if category_key in {"LOSSSIGNAL", "LOWSIGNAL"}:
                    ticket_type = "low_loss"
                elif category_key in {"ADDSPLITTER", "CHANGESPLITTER"}:
                    ticket_type = "change_add"
                else:
                    continue
            values = bucket(period_event.strftime("%Y-%m"), area)
            values["total"] += 1
            values["close_tickets" if is_closed else "open_tickets"] += 1
            values[f"tickets_{ticket_type}"] += 1
            assessed = not (ticket_type == "mp" and normalize(nro) == "NRO")
            if assessed and sla in {"SLA TERPENUHI", "MELEWATI SLA"}:
                result = "within" if sla == "SLA TERPENUHI" else "over"
                values[f"{ticket_type}_{result}"] += 1
                values[f"{result}_sla"] += 1
            opened_at, closed_at = as_datetime(opened), as_datetime(closed)
            if is_closed and assessed and opened_at and closed_at and closed_at > opened_at:
                values["mttr_seconds"] += (closed_at - opened_at).total_seconds()
                values["mttr_count"] += 1

    ordered_areas = list(area_names)
    months = []
    values = {}
    for key in sorted(counts, reverse=True):
        year, month = map(int, key.split("-"))
        days = monthrange(year, month)[1]
        months.append({"value": key, "label": f"{MONTH_NAMES[month]} {year}", "days": days})
        values[key] = {
            area: counts[key].get(area, bucket(key, area))
            for area in ordered_areas
        }
    open_list.sort(key=lambda item: item["opened"] or "")
    return {"months": months, "areas": ordered_areas, "values": values, "open_list": open_list}


def resolve_source() -> Path:
    global _current_source
    if _current_source and _current_source.is_file():
        return _current_source

    local = choose_newest_source(BASE_DIR.glob(f"*{SOURCE_LABEL}*.xlsx"))
    if local:
        _current_source = local
        return local

    for candidate in KNOWN_SOURCES:
        if candidate.is_file():
            _current_source = candidate
            return candidate

    discovered = []
    for root in SEARCH_ROOTS:
        if root.is_dir():
            discovered.extend(root.rglob(f"*{SOURCE_LABEL}*.xlsx"))
    source = choose_newest_source(discovered)
    if not source:
        raise FileNotFoundError(
            f"File Excel '{SOURCE_LABEL}' tidak ditemukan. "
            "Letakkan workbook di folder yang sama dengan dashboard, lalu klik Perbarui sekarang."
        )
    _current_source = source
    return source


def extract_dashboard(source: Path) -> dict[str, object]:
    if not source.is_file():
        raise FileNotFoundError(f"File Excel tidak ditemukan: {source}")

    stat = source.stat()
    workbook = load_workbook(source, read_only=True, data_only=True, keep_links=False)
    try:
        sheet = workbook["00. DASHBOARD"]
        data_sheet = workbook["Data Dashboard"]

        sla_by_area = {}
        for row in range(42, 58):
            area = sheet.cell(row, 2).value
            if not area:
                continue
            sla_by_area[str(area)] = {
                "sla_low_loss": number(sheet.cell(row, 3).value),
                "sla_change_add": number(sheet.cell(row, 5).value),
                "sla_mp": number(sheet.cell(row, 7).value),
                "sla_weighted": number(sheet.cell(row, 9).value),
                "status": json_value(sheet.cell(row, 11).value),
                "mttr": json_value(sheet.cell(row, 12).value),
            }

        overdue_by_area = {}
        for row in range(22, 38):
            area = sheet.cell(row, 18).value
            if not area:
                continue
            mp = number(sheet.cell(row, 20).value) or 0
            low_loss = number(sheet.cell(row, 21).value) or 0
            change_add = number(sheet.cell(row, 22).value) or 0
            overdue_by_area[str(area)] = {
                "rank": number(sheet.cell(row, 19).value),
                "mp": mp,
                "low_loss": low_loss,
                "change_add": change_add,
                "total": mp + low_loss + change_add,
            }

        areas = []
        for row in range(22, 38):
            area = sheet.cell(row, 2).value
            if not area:
                continue
            name = str(area)
            item = {
                "area": name,
                "total": number(sheet.cell(row, 3).value),
                "open": number(sheet.cell(row, 4).value),
                "close": number(sheet.cell(row, 5).value),
                "loss_signal": number(sheet.cell(row, 6).value),
                "low_signal": number(sheet.cell(row, 7).value),
                "add_splitter": number(sheet.cell(row, 8).value),
                "change_splitter": number(sheet.cell(row, 9).value),
                "mp": number(sheet.cell(row, 10).value),
                "within_sla": number(sheet.cell(row, 11).value),
                "over_sla": number(sheet.cell(row, 12).value),
                "sla_rate": number(sheet.cell(row, 13).value),
            }
            item.update(sla_by_area.get(name, {}))
            item["overdue"] = overdue_by_area.get(name, {})
            areas.append(item)

        trend = []
        for row in range(109, 116):
            day = sheet.cell(row, 8).value
            if not day:
                continue
            trend.append(
                {
                    "date": json_value(day),
                    "open": number(sheet.cell(row, 9).value) or 0,
                    "close": number(sheet.cell(row, 10).value) or 0,
                }
            )
        trend.reverse()

        target = number(sheet["M6"].value)
        payload = {
            "meta": {
                "source": str(source),
                "file_modified": datetime.fromtimestamp(stat.st_mtime).astimezone().isoformat(),
                "workbook_calculated": json_value(sheet["H19"].value),
                "served_at": datetime.now().astimezone().isoformat(),
                "refresh_seconds": 60,
            },
            "filters": {
                "area": json_value(sheet["B6"].value),
                "period": json_value(sheet["E6"].value),
                "year": json_value(sheet["H6"].value),
            },
            "settings": {
                "weight_mp": number(sheet["J6"].value),
                "weight_low_loss": number(sheet["K6"].value),
                "weight_change_add": number(sheet["L6"].value),
                "target_sla": target,
                "status_tolerance": number(data_sheet["V27"].value),
            },
            "metrics": {
                "total_tickets": number(sheet["B9"].value),
                "tickets_mp": number(sheet["E9"].value),
                "tickets_low_loss": number(sheet["H9"].value),
                "tickets_change_add": number(sheet["K9"].value),
                "sla_low_loss": number(sheet["B13"].value),
                "sla_change_add": number(sheet["E13"].value),
                "sla_mp": number(sheet["H13"].value),
                "sla_weighted": number(sheet["K13"].value),
                "open": number(sheet["B17"].value),
                "close": number(sheet["E17"].value),
                "open_today": number(sheet["H17"].value),
                "close_today": number(sheet["K17"].value),
                "within_sla": number(sheet["K38"].value),
                "over_sla": number(sheet["L38"].value),
                "sla_ticket_rate": number(sheet["M38"].value),
            },
            "areas": areas,
            "trend": trend,
            "daily": build_daily_report(workbook, source, [item["area"] for item in areas]),
        }
        return payload
    finally:
        workbook.close()


def dashboard_payload() -> tuple[dict[str, object], bool, str | None]:
    with _cache_lock:
        try:
            source = resolve_source()
            mtime_ns = source.stat().st_mtime_ns
            if (
                _cache["payload"] is None
                or _cache["source"] != source
                or _cache["mtime_ns"] != mtime_ns
            ):
                _cache["payload"] = extract_dashboard(source)
                _cache["source"] = source
                _cache["mtime_ns"] = mtime_ns
            payload = dict(_cache["payload"])
            payload["meta"] = dict(payload["meta"])
            payload["meta"]["served_at"] = datetime.now().astimezone().isoformat()
            return payload, False, None
        except Exception as exc:
            if _cache["payload"] is None:
                raise
            payload = dict(_cache["payload"])
            payload["meta"] = dict(payload["meta"])
            payload["meta"]["served_at"] = datetime.now().astimezone().isoformat()
            return payload, True, str(exc)


def port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as connection:
        connection.settimeout(0.2)
        return connection.connect_ex((HOST, port)) == 0


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def do_GET(self):
        if self.path.split("?", 1)[0] == "/api/dashboard":
            try:
                payload, stale, warning = dashboard_payload()
                payload["meta"]["stale"] = stale
                payload["meta"]["warning"] = warning
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
            except Exception as exc:
                body = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
                self.send_response(500)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()

    def log_message(self, format, *args):
        if self.path.startswith("/api/") and args and str(args[1]) == "200":
            return
        super().log_message(format, *args)


def main():
    server = None
    for port in range(START_PORT, START_PORT + 11):
        if port_in_use(port):
            continue
        try:
            server = ThreadingHTTPServer((HOST, port), DashboardHandler)
            break
        except OSError:
            continue
    if server is None:
        raise OSError(f"Port {START_PORT}-{START_PORT + 10} sedang digunakan.")

    url = f"http://{HOST}:{server.server_port}/"
    print(f"Dashboard aktif: {url}")
    try:
        print(f"Sumber Excel: {resolve_source()}")
    except FileNotFoundError as exc:
        print(exc)
    print("Tekan Ctrl+C untuk menghentikan.")
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
