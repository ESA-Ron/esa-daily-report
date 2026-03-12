import os
import io
import anthropic
import openpyxl
import csv
import traceback
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

def excel_to_csv_text(file_bytes, max_rows=5000):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    output = io.StringIO()
    writer = csv.writer(output)
    count = 0
    header_found = False
    for row in ws.iter_rows(values_only=True):
        if count >= max_rows:
            break
        non_empty = [c for c in row if c is not None and str(c).strip() != '']
        if not header_found:
            if len(non_empty) >= 3:
                header_found = True
            else:
                continue
        writer.writerow([str(c) if c is not None else '' for c in row])
        count += 1
    wb.close()
    return output.getvalue(), count

@app.route('/run-report', methods=['POST'])
def run_report():
    try:
        api_key = os.environ.get('ANTHROPIC_API_KEY') or request.headers.get('X-Api-Key', '')
        if not api_key or not api_key.startswith('sk-ant-'):
            return jsonify({'error': 'Invalid or missing API key.'}), 400

        esa_file = request.files.get('esa_file')
        map_file = request.files.get('map_file')
        if not esa_file:
            return jsonify({'error': 'No ESA file uploaded.'}), 400

        esa_bytes = esa_file.read()
        map_bytes = map_file.read() if map_file else None

        esa_csv, esa_rows = excel_to_csv_text(esa_bytes, max_rows=5000)
        map_csv, _ = excel_to_csv_text(map_bytes, max_rows=1400) if map_bytes else ('', 0)
        has_mapping = bool(map_csv)

        # Hard trim
        esa_trimmed = esa_csv[:60000]
        map_trimmed = map_csv[:12000]

        leaderboard_section = (
            "8. SALESPERSON LEADERBOARD: Use mapping CSV. Join Sold-To No = Southwire # column. "
            "Status must = Active. Exclude blank or 'Need' names. "
            "Show Top 5 Outside Salesperson, Top 5 Inside Salesperson, Top 3 Sales Teams — all by net sales."
        ) if has_mapping else (
            "8. SALESPERSON LEADERBOARD: No mapping file. Show placeholder."
        )

        prompt = f"""You are a financial analyst for ESA (Electrical Sales Associates).
Analyze the ESA Bookings CSV and return a complete HTML daily report.

COLUMNS: Sales Order No (distinct count=orders), Sold-To No, Sold-To Name, Created On Dt., EST Ship Date, Net Sales (USD), Net weight, Net weight Unit (use LB only), Material Text, PH1.

BUILD THESE 8 SECTIONS WITH REAL NUMBERS:
1. MTD SUMMARY: distinct orders, total net sales, total LB pounds, plan $25M, avg/day stats, date range
2. FORECAST VS PLAN: EOM run-rate forecast, variance vs $25M in $ and %
3. SHIPPING OUTLOOK: avg/median lead days, avg/median EST ship date  
4. DAILY TABLE: date, orders, sales, pounds, cumulative columns
5. TOP 5 CUSTOMERS: by net sales, with order count and pounds
6. TOP 5 MATERIALS: by net sales with pounds
7. BUSINESS MIX: top 6 PH1 codes by sales
{leaderboard_section}

HTML: self-contained, navy #1F4E79 header, white body, KPI cards, styled tables, Arial font.
Format: $1,234,567.89 and 1,234,567 LB.
RETURN ONLY HTML starting with <!DOCTYPE html>. No markdown.

=== ESA DATA ({esa_rows} rows) ===
{esa_trimmed}
"""
        if has_mapping:
            prompt += f"\n=== MAPPING DATA ===\n{map_trimmed}"

        client = anthropic.Anthropic(api_key=api_key, timeout=120.0)
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            messages=[{"role": "user", "content": prompt}]
        )

        html = message.content[0].text.replace('```html','').replace('```','').strip()
        idx = html.find('<!DOCTYPE')
        if idx > 0:
            html = html[idx:]

        return jsonify({'html': html, 'success': True})

    except anthropic.AuthenticationError:
        return jsonify({'error': 'API key invalid. Check console.anthropic.com.'}), 401
    except anthropic.RateLimitError:
        return jsonify({'error': 'Rate limit hit. Wait 30 seconds and try again.'}), 429
    except anthropic.APITimeoutError:
        return jsonify({'error': 'Request timed out. Try again — Claude may be busy.'}), 504
    except Exception as e:
        tb = traceback.format_exc()
        print(f"ERROR: {tb}")
        return jsonify({'error': f'{type(e).__name__}: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
