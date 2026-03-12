import os
import io
import anthropic
import openpyxl
import csv
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

def excel_to_csv_text(file_bytes, max_rows=8000):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    output = io.StringIO()
    writer = csv.writer(output)
    count = 0
    for row in ws.iter_rows(values_only=True):
        if count >= max_rows:
            break
        if any(cell is not None for cell in row):
            writer.writerow([str(c) if c is not None else '' for c in row])
            count += 1
    wb.close()
    return output.getvalue()

@app.route('/run-report', methods=['POST'])
def run_report():
    try:
        api_key = os.environ.get('ANTHROPIC_API_KEY') or request.headers.get('X-Api-Key', '')
        if not api_key or not api_key.startswith('sk-ant-'):
            return jsonify({'error': 'Invalid or missing API key. Check your key and try again.'}), 400

        esa_file = request.files.get('esa_file')
        map_file = request.files.get('map_file')

        if not esa_file:
            return jsonify({'error': 'No ESA file uploaded.'}), 400

        esa_csv = excel_to_csv_text(esa_file.read())
        map_csv = excel_to_csv_text(map_file.read()) if map_file else None
        has_mapping = map_csv is not None

        leaderboard_instruction = """
8. SALESPERSON LEADERBOARD: Use the mapping CSV to join on Sold-To No = Southwire #.
   Only use rows where Status = Active and salesperson is not blank or 'Need'.
   Show: Top 5 Outside Salesperson by net sales. Top 5 Inside Salesperson by net sales. Top 3 Sales Teams by net sales.
""" if has_mapping else """
8. SALESPERSON LEADERBOARD: No mapping file provided. Show placeholder and list top 5 customers by Sold-To Name as proxy.
"""

        prompt = f"""You are a senior financial analyst for ESA (Electrical Sales Associates).
Analyze the ESA Bookings Report CSV data below and produce a complete professional HTML daily report email.

Header row is around row 3. Key columns: Sales Order No, Sold-To No, Sold-To Name, Created On Dt.,
EST Ship Date, Net Sales (USD), Net weight, Net weight Unit, Material Text, PH1, PH2.

REPORT SECTIONS (ALL with REAL calculated numbers):

1. MTD SUMMARY: Distinct Sales Order No count = total orders. Sum Net Sales (USD). Sum Net weight where unit=LB.
   Monthly plan = $25,000,000. Show avg orders/day, avg sales/day, date range.

2. FORECAST VS PLAN: EOM forecast = (MTD sales / elapsed days) x days in month. Variance vs $25M.

3. SHIPPING OUTLOOK: Per distinct order: lead days = EST Ship Date minus Created On Dt.
   Show avg lead days, median lead days, avg EST Ship Date, median EST Ship Date.

4. DAILY TABLE: Group by date. Orders, net sales, net pounds, cumulative totals.

5. TOP 5 CUSTOMERS: By net sales. Name, sales, order count, pounds.

6. TOP 5 MATERIALS: By net sales from Material Text. Name, sales, pounds.

7. BUSINESS MIX: Group by PH1. Code, sales, order count. Top 6.

{leaderboard_instruction}

DESIGN: Self-contained HTML, inline CSS. Navy header #1F4E79, white body, KPI cards, alternating table rows, Arial font.
Dollars as $1,234,567.89 — Pounds as 1,234,567 LB.
Return ONLY complete HTML starting with <!DOCTYPE html>. No markdown.

=== ESA BOOKINGS CSV ===
{esa_csv[:80000]}
"""
        if has_mapping:
            prompt += f"\n=== MAPPING CSV ===\n{map_csv[:20000]}"

        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            messages=[{"role": "user", "content": prompt}]
        )

        html = message.content[0].text.replace('```html', '').replace('```', '').strip()
        return jsonify({'html': html, 'success': True})

    except anthropic.AuthenticationError:
        return jsonify({'error': 'API key invalid. Go to console.anthropic.com and create a new key.'}), 401
    except anthropic.RateLimitError:
        return jsonify({'error': 'Rate limit hit. Wait 30 seconds and try again.'}), 429
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
