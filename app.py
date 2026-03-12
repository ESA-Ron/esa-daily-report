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

def excel_to_csv_text(file_bytes, max_rows=6000, skip_empty_leading=True):
    """Convert Excel bytes to CSV string, finding the real header row."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    output = io.StringIO()
    writer = csv.writer(output)
    count = 0
    header_found = False

    for row in ws.iter_rows(values_only=True):
        if count >= max_rows:
            break
        # Skip leading empty rows before header
        if skip_empty_leading and not header_found:
            non_empty = [c for c in row if c is not None and str(c).strip() != '']
            if len(non_empty) < 3:
                continue
            header_found = True
        writer.writerow([str(c) if c is not None else '' for c in row])
        count += 1

    wb.close()
    return output.getvalue()

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

        # Read files into memory first
        esa_bytes = esa_file.read()
        map_bytes = map_file.read() if map_file else None

        esa_csv = excel_to_csv_text(esa_bytes, max_rows=6000)
        map_csv = excel_to_csv_text(map_bytes, max_rows=1500, skip_empty_leading=False) if map_bytes else None
        has_mapping = map_csv is not None

        # Trim to safe token limits
        esa_trimmed = esa_csv[:70000]
        map_trimmed = map_csv[:15000] if map_csv else ''

        leaderboard_section = (
            "8. SALESPERSON LEADERBOARD: Use the mapping CSV. Join on Sold-To No = Southwire # column. "
            "Filter Status = Active only, exclude blank or 'Need' salesperson names. "
            "Show: Top 5 Outside Salesperson by net sales MTD. Top 5 Inside Salesperson by net sales MTD. "
            "Top 3 Sales Teams by net sales MTD."
        ) if has_mapping else (
            "8. SALESPERSON LEADERBOARD: No mapping file provided. Show placeholder message."
        )

        prompt = f"""You are a senior financial analyst for ESA (Electrical Sales Associates).
Analyze the ESA Bookings Report data (CSV format) and build a complete professional HTML daily report.

KEY COLUMNS in the ESA data:
- Sales Order No → count distinct values for total orders
- Sold-To No, Sold-To Name → customer identifier
- Created On Dt. → order creation date (format: M/D/YYYY or YYYY-MM-DD)
- EST Ship Date → estimated ship date
- Net Sales (USD) → revenue amount
- Net weight → pounds value
- Net weight Unit → only include rows where this = LB
- Material Text → product/material name
- PH1 → business mix category code

REQUIRED SECTIONS (calculate ALL from the actual data):

1. MTD SUMMARY
   - Total distinct Sales Order No = order count
   - Sum of Net Sales (USD) = total revenue
   - Sum of Net weight where Net weight Unit = LB = total pounds
   - Monthly plan: $25,000,000
   - Avg orders per day, avg revenue per day
   - Date range (min to max Created On Dt.)

2. FORECAST VS PLAN
   - Days elapsed = (max date - first date of month + 1)
   - Days in month = total calendar days in the month
   - EOM forecast = (MTD revenue / days elapsed) x days in month
   - $ variance vs $25M plan, % variance

3. SHIPPING OUTLOOK
   - For each distinct Sales Order No, lead days = EST Ship Date minus Created On Dt.
   - Average lead days (round to 1 decimal)
   - Median lead days
   - Average EST Ship Date
   - Median EST Ship Date

4. DAILY BREAKDOWN TABLE
   - Group by Created On Dt. date
   - Columns: Date | Orders | Net Sales | Net Pounds | Cumul. Orders | Cumul. Sales | Cumul. Pounds

5. TOP 5 CUSTOMERS by Net Sales
   - Sold-To Name, Net Sales USD, Order Count, Net Pounds LB

6. TOP 5 MATERIALS by Net Sales
   - Material Text, Net Sales USD, Net Pounds LB

7. BUSINESS MIX (top 6 PH1 codes)
   - PH1 code, Net Sales USD, Order Count

{leaderboard_section}

HTML DESIGN REQUIREMENTS:
- Complete self-contained HTML with all CSS in <style> tag
- Header: dark navy #1F4E79 background, white text, ESA Daily Report title
- Body: white background, clean sans-serif font (Arial)
- KPI Summary: 3-column card grid with colored left borders (blue, green, gold)
- All tables: alternating row colors (#f8fafc), header row in #1F4E79 white text
- Dollar amounts: $1,234,567.89 format
- Pound amounts: 1,234,567 LB format
- Mobile responsive

CRITICAL: Return ONLY the complete HTML document. Start with <!DOCTYPE html>. No markdown fences, no explanation text.

=== ESA BOOKINGS DATA (CSV) ===
{esa_trimmed}
"""
        if has_mapping:
            prompt += f"\n=== CUSTOMER TO SALESPERSON MAPPING (CSV) ===\n{map_trimmed}"

        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            messages=[{"role": "user", "content": prompt}]
        )

        html = message.content[0].text.replace('```html', '').replace('```', '').strip()
        if not html.startswith('<!'):
            # Find the start of HTML if Claude added any preamble
            idx = html.find('<!DOCTYPE')
            if idx > 0:
                html = html[idx:]

        return jsonify({'html': html, 'success': True})

    except anthropic.AuthenticationError:
        return jsonify({'error': 'API key invalid. Go to console.anthropic.com and create a new key.'}), 401
    except anthropic.RateLimitError:
        return jsonify({'error': 'Rate limit hit. Wait 30 seconds and try again.'}), 429
    except MemoryError:
        return jsonify({'error': 'File too large for free server. Try uploading just the ESA file without the mapping file first.'}), 500
    except Exception as e:
        return jsonify({'error': f'{type(e).__name__}: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
