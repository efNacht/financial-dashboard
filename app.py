"""
Financial Dashboard — Flask web application.
Parses Russian financial reports (БФО) and displays interactive analytics.
"""

import os
import json
import sqlite3
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from parser import parse_xlsx, detect_company_year, compute_metrics, get_company_info

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

APP_DIR = Path(__file__).resolve().parent
DB_PATH = os.environ.get('DB_PATH', str(APP_DIR / 'data.db'))
UPLOAD_DIR = Path(os.environ.get('UPLOAD_DIR', str(APP_DIR / 'uploads')))
UPLOAD_DIR.mkdir(exist_ok=True)


# ============================================================
# DATABASE
# ============================================================

def get_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db


def init_db():
    db = get_db()
    db.execute('''CREATE TABLE IF NOT EXISTS companies (
        short_code TEXT PRIMARY KEY,
        full_name TEXT NOT NULL,
        inn TEXT DEFAULT ''
    )''')
    db.execute('''CREATE TABLE IF NOT EXISTS reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        short_code TEXT NOT NULL,
        year INTEGER NOT NULL,
        metrics TEXT NOT NULL,
        raw_pl TEXT DEFAULT '{}',
        raw_balance TEXT DEFAULT '{}',
        filename TEXT DEFAULT '',
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(short_code, year),
        FOREIGN KEY (short_code) REFERENCES companies(short_code)
    )''')
    db.commit()
    db.close()


def get_all_data():
    """Get all company data formatted for the dashboard JS."""
    db = get_db()
    companies = {r['short_code']: {'full_name': r['full_name'], 'inn': r['inn']}
                 for r in db.execute('SELECT * FROM companies').fetchall()}

    reports = db.execute('SELECT * FROM reports ORDER BY short_code, year').fetchall()
    db.close()

    result = {}
    for r in reports:
        sc = r['short_code']
        if sc not in result:
            info = companies.get(sc, {'full_name': sc, 'inn': ''})
            result[sc] = {
                'full_name': info['full_name'],
                'inn': info['inn'],
                'short_name': sc,
                'years_data': {},
            }
        result[sc]['years_data'][r['year']] = json.loads(r['metrics'])

    # Sort by latest revenue descending
    def sort_key(item):
        years = item[1]['years_data']
        if not years:
            return 0
        latest = max(years.keys())
        return -(years[latest].get('revenue', 0) or 0)

    result = dict(sorted(result.items(), key=sort_key))
    return result


def save_report(short_code, year, metrics, raw_pl, raw_balance, filename):
    """Save or update a parsed report in the database."""
    full_name, inn = get_company_info(short_code)
    db = get_db()
    db.execute('INSERT OR REPLACE INTO companies (short_code, full_name, inn) VALUES (?, ?, ?)',
               (short_code, full_name, inn))
    db.execute('''INSERT OR REPLACE INTO reports (short_code, year, metrics, raw_pl, raw_balance, filename)
                  VALUES (?, ?, ?, ?, ?, ?)''',
               (short_code, year, json.dumps(metrics, ensure_ascii=False),
                json.dumps(raw_pl, ensure_ascii=False), json.dumps(raw_balance, ensure_ascii=False),
                filename))
    db.commit()
    db.close()


# ============================================================
# ROUTES
# ============================================================

@app.route('/')
def index():
    return redirect(url_for('dashboard'))


@app.route('/dashboard')
def dashboard():
    data = get_all_data()
    return render_template('dashboard.html', data_json=json.dumps(data, ensure_ascii=False))


@app.route('/summary')
def summary():
    data = get_all_data()
    return render_template('summary.html', data_json=json.dumps(data, ensure_ascii=False))


@app.route('/upload', methods=['GET'])
def upload_page():
    db = get_db()
    reports = db.execute('''SELECT c.full_name, r.short_code, r.year, r.filename, r.uploaded_at
                           FROM reports r JOIN companies c ON r.short_code = c.short_code
                           ORDER BY r.uploaded_at DESC''').fetchall()
    companies = db.execute('SELECT short_code, full_name FROM companies ORDER BY full_name').fetchall()
    db.close()
    return render_template('upload.html', reports=reports, companies=companies)


@app.route('/api/upload', methods=['POST'])
def api_upload():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'Файлы не выбраны'}), 400

    results = []
    for f in files:
        if not f.filename or not f.filename.endswith('.xlsx'):
            results.append({'file': f.filename, 'status': 'error', 'message': 'Не xlsx файл'})
            continue

        short_code, year = detect_company_year(f.filename)
        if not short_code or not year:
            results.append({'file': f.filename, 'status': 'error',
                           'message': 'Не удалось определить компанию/год из имени файла. Формат: КОМПАНИЯ_ГОД.xlsx'})
            continue

        # Save file temporarily
        filepath = UPLOAD_DIR / f.filename
        f.save(filepath)

        try:
            raw = parse_xlsx(str(filepath))
            metrics = compute_metrics(raw.get('pl', {}), raw.get('balance', {}))
            save_report(short_code, year, metrics, raw.get('pl', {}), raw.get('balance', {}), f.filename)

            full_name, _ = get_company_info(short_code)
            results.append({
                'file': f.filename,
                'status': 'ok',
                'company': full_name,
                'short_code': short_code,
                'year': year,
                'revenue': metrics.get('revenue', 0),
                'net_profit': metrics.get('net_profit', 0),
            })
        except Exception as e:
            results.append({'file': f.filename, 'status': 'error', 'message': str(e)})
        finally:
            filepath.unlink(missing_ok=True)

    return jsonify({'results': results})


@app.route('/api/delete', methods=['POST'])
def api_delete():
    data = request.json
    short_code = data.get('short_code')
    year = data.get('year')
    db = get_db()
    if year:
        db.execute('DELETE FROM reports WHERE short_code = ? AND year = ?', (short_code, int(year)))
    else:
        db.execute('DELETE FROM reports WHERE short_code = ?', (short_code,))
        db.execute('DELETE FROM companies WHERE short_code = ?', (short_code,))
    db.commit()
    db.close()
    return jsonify({'ok': True})


@app.route('/api/data')
def api_data():
    return jsonify(get_all_data())


@app.route('/export/xlsx')
def export_xlsx():
    """Export summary table as xlsx."""
    data = get_all_data()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Сводная таблица"

    # Styles
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    num_fmt = '#,##0'
    pct_fmt = '0%'

    headers = [
        "Юр. лицо", "Выручка 2023", "Выручка 2024", "Выручка 2025",
        "Чистая прибыль 2023", "Чистая прибыль 2024", "Чистая прибыль 2025",
        "Прибыль 24/23, %", "Прибыль 25/24, %", "Прибыль 25−24, руб.", "Прибыль за 2 года"
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    row_num = 2
    for sc, company in data.items():
        yd = company['years_data']
        d23 = yd.get(2023, yd.get('2023', {}))
        d24 = yd.get(2024, yd.get('2024', {}))
        d25 = yd.get(2025, yd.get('2025', {}))

        r23 = d23.get('revenue', 0) or 0
        r24 = d24.get('revenue', 0) or 0
        r25 = d25.get('revenue', 0) or 0
        p23 = d23.get('net_profit', 0) or 0
        p24 = d24.get('net_profit', 0) or 0
        p25 = d25.get('net_profit', 0) or 0

        g2423 = (p24 - p23) / abs(p23) if p23 else None
        g2524 = (p25 - p24) / abs(p24) if p24 else None
        diff = p25 - p24
        sum2y = p24 + p25

        values = [
            company['full_name'],
            r23 or None, r24 or None, r25 or None,
            p23 or None, p24 or None, p25 or None,
            g2423, g2524, diff, sum2y,
        ]

        for col, v in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col, value=v)
            cell.border = thin_border
            if col == 1:
                cell.alignment = Alignment(horizontal="left")
            elif col in (8, 9):
                cell.number_format = '0%'
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.number_format = num_fmt
                cell.alignment = Alignment(horizontal="right")

            # Color negative values red
            if isinstance(v, (int, float)) and v < 0:
                cell.font = Font(color="DC2626")

        row_num += 1

    # Auto-width
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 18
    ws.column_dimensions['A'].width = 35

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, download_name="сводная_таблица.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================
# SEED DATA — load existing xlsx files on first run
# ============================================================

def seed_from_directory(directory):
    """Parse all xlsx files from a directory and save to DB."""
    d = Path(directory)
    if not d.exists():
        return 0
    count = 0
    for fp in sorted(d.glob('*.xlsx')):
        short_code, year = detect_company_year(fp.name)
        if not short_code or not year:
            continue
        try:
            raw = parse_xlsx(str(fp))
            metrics = compute_metrics(raw.get('pl', {}), raw.get('balance', {}))
            save_report(short_code, year, metrics, raw.get('pl', {}), raw.get('balance', {}), fp.name)
            count += 1
        except Exception as e:
            print(f"  Error parsing {fp.name}: {e}")
    return count


# ============================================================
# INIT — runs on import (gunicorn, preview, or __main__)
# ============================================================

init_db()

# Seed from local directory if DB is empty
_db = get_db()
_cnt = _db.execute('SELECT COUNT(*) as c FROM reports').fetchone()['c']
_db.close()
if _cnt == 0:
    _reports_dir = os.environ.get('REPORTS_DIR', '')
    if _reports_dir and Path(_reports_dir).exists():
        print(f"Seeding from {_reports_dir}...")
        _n = seed_from_directory(_reports_dir)
        print(f"Loaded {_n} reports")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=os.environ.get('FLASK_DEBUG', '0') == '1')
