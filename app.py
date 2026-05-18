"""Sığorta Hesabatı Generatoru — Flask web app.

Upload a hospital services .xlsx, group rows by column N (Müəssisə), and
download per-insurance reports containing only the yellow columns.
"""
from __future__ import annotations

import io
import os
import re
import uuid
import zipfile
from datetime import datetime, date
from threading import Lock

import pandas as pd
from flask import (
    Flask,
    Response,
    abort,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024  # 250 MB cap

# Set APP_PASSWORD in the environment to require a password to access the app.
# If unset, the app is open (useful for local development only).
APP_PASSWORD = os.environ.get("APP_PASSWORD", "").strip() or None


@app.before_request
def _require_password():
    if APP_PASSWORD is None:
        return None
    auth = request.authorization
    if auth is not None and auth.password == APP_PASSWORD:
        return None
    return Response(
        "Daxil olmaq üçün parol tələb olunur.\n",
        401,
        {"WWW-Authenticate": 'Basic realm="Sigorta Hesabati"'},
    )

# Source-column positions (0-indexed) we keep in the output, in display order.
# We never rename these — the source file's headers are preserved verbatim.
KEEP_COL_POSITIONS = [
    0,   # A  Xidmət tarixi
    3,   # D  Protokol No
    4,   # E  Soyadı
    5,   # F  Adı
    6,   # G  Baba Adı
    7,   # H  Tam adi
    12,  # M  Doğum Tarixi
    32,  # AG Xidmət adı
    34,  # AI Ədəd
    40,  # AO Kdvli Hasta Tutar
]
GROUP_COL_IDX = 13  # column N: Müəssisə
ALL_COL_IDX = sorted({*KEEP_COL_POSITIONS, GROUP_COL_IDX})

# Column-type classification by source position (not header label, so we can
# preserve headers exactly as they appear in the source file).
DATE_COL_POSITIONS = {0, 12}     # Xidmət tarixi, Doğum Tarixi
MONEY_COL_POSITIONS = {40}        # Kdvli Hasta Tutar
QTY_COL_POSITIONS = {34}          # Ədəd

UNKNOWN_LABEL = "Bilinmir"  # used when Müəssisə is empty/NaN

# In-memory store keyed by upload id:
#   {upload_id: {"filename": str, "groups": {name: DataFrame},
#                "col_types": {col_name: 'date'|'money'|'qty'|'text'},
#                "created": datetime}}
STORE: dict[str, dict] = {}
STORE_LOCK = Lock()


# ---------- helpers ----------


def _read_and_group(file_storage) -> tuple[dict[str, pd.DataFrame], dict[str, str]]:
    """Read the uploaded xlsx and return (groups, col_types).

    Source-file column headers are preserved verbatim.
    """
    df = pd.read_excel(file_storage, usecols=ALL_COL_IDX, engine="openpyxl")

    # Map source positions -> the header pandas read for that column.
    pos_to_name: dict[int, str] = {idx: df.columns[i] for i, idx in enumerate(ALL_COL_IDX)}
    group_col = pos_to_name[GROUP_COL_IDX]
    keep_col_names = [pos_to_name[p] for p in KEEP_COL_POSITIONS]

    # Classify each kept column by its source position so the format helpers
    # know what to do without relying on the (preserved) header text.
    col_types: dict[str, str] = {}
    for pos in KEEP_COL_POSITIONS:
        name = pos_to_name[pos]
        if pos in DATE_COL_POSITIONS:
            col_types[name] = "date"
        elif pos in MONEY_COL_POSITIONS:
            col_types[name] = "money"
        elif pos in QTY_COL_POSITIONS:
            col_types[name] = "qty"
        else:
            col_types[name] = "text"

    groups: dict[str, pd.DataFrame] = {}
    for name, sub in df.groupby(group_col, dropna=False, sort=True):
        if pd.isna(name) or str(name).strip() == "":
            key = UNKNOWN_LABEL
        else:
            key = str(name).strip()
        sub = sub[keep_col_names].reset_index(drop=True)
        if key in groups:
            groups[key] = pd.concat([groups[key], sub], ignore_index=True)
        else:
            groups[key] = sub
    return groups, col_types


_DATE_PATTERNS = [
    "%Y-%m-%d %H:%M:%S.%f",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y",
    "%d.%m.%Y",
]


def _try_parse_date(s: str):
    s = s.strip()
    if not s:
        return None
    for fmt in _DATE_PATTERNS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _fmt_cell(col_type: str, val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if pd.isna(val):
        return ""
    if col_type == "date":
        if isinstance(val, (pd.Timestamp, datetime, date)):
            return val.strftime("%Y-%m-%d")
        s = str(val).strip()
        if " " in s and re.match(r"^\d{4}-\d{2}-\d{2}", s):
            s = s.split(" ", 1)[0]
        return s
    if col_type == "money":
        try:
            return f"{float(val):,.2f}"
        except (TypeError, ValueError):
            return str(val)
    if col_type == "qty":
        try:
            f = float(val)
            return str(int(f)) if f.is_integer() else f"{f:g}"
        except (TypeError, ValueError):
            return str(val)
    return str(val)


def _df_to_xlsx_bytes(df: pd.DataFrame, col_types: dict[str, str], sheet_name: str = "Sheet1") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = _safe_sheet_name(sheet_name, set())
    _write_df_to_sheet(ws, df, col_types)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_df_to_sheet(ws, df: pd.DataFrame, col_types: dict[str, str]) -> None:
    headers = list(df.columns)
    ws.append(headers)
    header_font = Font(bold=True, color="000000")
    header_fill = PatternFill("solid", fgColor="FFF2A8")  # soft yellow
    for col_idx, _ in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col_idx)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")

    for row in df.itertuples(index=False, name=None):
        out = []
        for col_name, val in zip(headers, row):
            ctype = col_types.get(col_name, "text")
            if val is None or (isinstance(val, float) and pd.isna(val)):
                out.append(None)
            elif ctype == "date":
                if isinstance(val, pd.Timestamp):
                    out.append(val.to_pydatetime())
                elif isinstance(val, (datetime, date)):
                    out.append(val)
                else:
                    parsed = _try_parse_date(str(val))
                    out.append(parsed if parsed is not None else val)
            else:
                out.append(val)
        ws.append(out)

    money_fmt = "#,##0.00"
    date_fmt = "yyyy-mm-dd"
    n_rows = ws.max_row
    for col_idx, name in enumerate(headers, start=1):
        ctype = col_types.get(name, "text")
        if ctype == "date":
            for r in range(2, n_rows + 1):
                ws.cell(row=r, column=col_idx).number_format = date_fmt
        elif ctype == "money":
            for r in range(2, n_rows + 1):
                ws.cell(row=r, column=col_idx).number_format = money_fmt

    # Auto-fit-ish column widths (sample first 200 rows; cap to keep sane).
    for col_idx, name in enumerate(headers, start=1):
        max_len = len(str(name))
        for r in range(2, min(n_rows, 201) + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None:
                continue
            s = v.strftime("%Y-%m-%d") if isinstance(v, (datetime, date)) else str(v)
            if len(s) > max_len:
                max_len = len(s)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 45)

    ws.freeze_panes = "A2"


def _safe_sheet_name(name: str, used: set[str]) -> str:
    s = re.sub(r"[\[\]\\\/\?\*\:]", "_", str(name)).strip()[:31] or "Sheet"
    base = s
    i = 1
    while s in used:
        suffix = f"_{i}"
        s = (base[: 31 - len(suffix)]) + suffix
        i += 1
    used.add(s)
    return s


def _safe_filename(name: str) -> str:
    """Filesystem-safe filename component; preserves Azerbaijani/Turkish letters."""
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(name)).strip().strip(".")
    return s[:120] or "sigorta"


def _get_upload(upload_id: str) -> dict:
    with STORE_LOCK:
        rec = STORE.get(upload_id)
    if not rec:
        abort(404, description="Yükləmə tapılmadı və ya vaxtı keçib. Zəhmət olmasa yenidən yükləyin.")
    return rec


# ---------- routes ----------


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "Fayl təqdim edilmədi"}), 400
    fname = secure_filename(f.filename) or "upload.xlsx"
    if not fname.lower().endswith(".xlsx"):
        return jsonify({"error": "Zəhmət olmasa .xlsx faylı yükləyin"}), 400

    try:
        groups, col_types = _read_and_group(f)
    except Exception as e:
        return jsonify({"error": f"Fayl oxuna bilmədi: {e}"}), 400

    upload_id = uuid.uuid4().hex
    with STORE_LOCK:
        STORE[upload_id] = {
            "filename": fname,
            "groups": groups,
            "col_types": col_types,
            "created": datetime.utcnow(),
        }

    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify({"upload_id": upload_id, "redirect": url_for("results", upload_id=upload_id)})
    return redirect(url_for("results", upload_id=upload_id))


@app.route("/results/<upload_id>")
def results(upload_id):
    rec = _get_upload(upload_id)
    groups = rec["groups"]
    items = [
        {"name": name, "rows": len(df)}
        for name, df in sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    ]
    total_rows = sum(len(df) for df in groups.values())
    return render_template(
        "results.html",
        upload_id=upload_id,
        filename=rec["filename"],
        items=items,
        total_rows=total_rows,
        total_groups=len(items),
    )


@app.route("/preview/<upload_id>/<path:name>")
def preview(upload_id, name):
    rec = _get_upload(upload_id)
    groups = rec["groups"]
    col_types = rec["col_types"]
    if name not in groups:
        abort(404, description=f"'{name}' sığortası tapılmadı.")
    df = groups[name]
    headers = list(df.columns)

    max_preview = 2000
    total = len(df)
    shown = min(total, max_preview)
    rows = []
    for tup in df.head(shown).itertuples(index=False, name=None):
        rows.append([_fmt_cell(col_types.get(headers[i], "text"), v) for i, v in enumerate(tup)])

    return render_template(
        "preview.html",
        upload_id=upload_id,
        name=name,
        headers=headers,
        rows=rows,
        total=total,
        shown=shown,
        truncated=total > shown,
    )


@app.route("/download/<upload_id>/<path:name>")
def download_one(upload_id, name):
    rec = _get_upload(upload_id)
    groups = rec["groups"]
    if name not in groups:
        abort(404)
    data = _df_to_xlsx_bytes(groups[name], rec["col_types"], sheet_name=name)
    fname = f"{_safe_filename(name)}.xlsx"
    return send_file(
        io.BytesIO(data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=fname,
    )


@app.route("/download_all/<upload_id>")
def download_all(upload_id):
    """ZIP archive: one separate .xlsx file per insurance."""
    rec = _get_upload(upload_id)
    groups = rec["groups"]
    col_types = rec["col_types"]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names: set[str] = set()
        for name, df in groups.items():
            base = _safe_filename(name)
            arc = f"{base}.xlsx"
            k = 1
            while arc in used_names:
                arc = f"{base}_{k}.xlsx"
                k += 1
            used_names.add(arc)
            zf.writestr(arc, _df_to_xlsx_bytes(df, col_types, sheet_name=name))
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="sigorta_hesabatlari.zip",
    )


@app.route("/clear/<upload_id>", methods=["POST"])
def clear(upload_id):
    with STORE_LOCK:
        STORE.pop(upload_id, None)
    return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    # Bind on 0.0.0.0 so it works under Render's port forwarding; on local
    # development this is harmless because the cloud platform isn't listening.
    host = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
    app.run(host=host, port=port, debug=debug)
