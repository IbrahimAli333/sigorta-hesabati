# Sığorta Hesabatı Generatoru

Localhost Flask app that takes a hospital services `.xlsx`, groups rows by column **N (Müəssisə)**, and produces a per-insurance report containing only these source columns (headers preserved verbatim):

| Col | Header                |
|-----|-----------------------|
| A   | Xidmət tarixi         |
| D   | Protokol No           |
| E   | Soyadı                |
| F   | Adı                   |
| G   | Baba Adı              |
| H   | Tam adi               |
| M   | Doğum Tarixi          |
| AG  | Xidmət adı            |
| AI  | Ədəd                  |
| AO  | Kdvli Hasta Tutar     |

## Run

```bash
cd ~/saf_hospital/flask_app
pip3 install -r requirements.txt
python3 app.py
```

Open <http://127.0.0.1:5050>.

## Features

- Drag-and-drop / browse upload (up to 250 MB, ~63k+ rows tested)
- All UI in Azerbaijani
- List of every insurance with row counts and a filter box
- In-browser preview per insurance (capped at 2 000 rows for responsiveness; the downloaded file always contains everything)
- **Per-insurance download** — one `.xlsx` per click
- **Hamısını yüklə (.zip)** — every insurance as a separate `.xlsx` bundled into one ZIP archive (`sigorta_hesabatlari.zip`)
- UTF-8 throughout: Azerbaijani/Turkish letters survive in column headers, cell content, sheet names, and ZIP file names
- Date strings like `"2003-01-18 00:00:00.0"` are parsed back into real Excel dates with `yyyy-mm-dd` format; the money column uses `#,##0.00`
- Source-file column headers are preserved verbatim (no renaming)

## Notes

- Uploaded data lives in memory keyed by an upload ID and is cleared on server restart.
- Sheet names sanitized for Excel (31-char limit, no `[]\/?*:`).
