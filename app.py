import os
import base64
import anthropic
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__, static_folder='static')
CORS(app)

# ── Serve the frontend ──────────────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

# ── Run the ESA report ──────────────────────────────────────────────────────
@app.route('/run-report', methods=['POST'])
def run_report():
    try:
        # Get API key — from env var (Render secret) or from request header
        api_key = os.environ.get('ANTHROPIC_API_KEY') or request.headers.get('X-Api-Key', '')
        if not api_key or not api_key.startswith('sk-ant-'):
            return jsonify({'error': 'Invalid or missing API key. Check your key and try again.'}), 400

        # Get uploaded files
        esa_file = request.files.get('esa_file')
        map_file = request.files.get('map_file')

        if not esa_file:
            return jsonify({'error': 'No ESA file uploaded.'}), 400

        esa_b64 = base64.b64encode(esa_file.read()).decode('utf-8')

        # Build message content
        content = [
            {
                "type": "text",
                "text": """You are a senior financial analyst for ESA (Electrical Sales Associates).
Analyze the attached ESA Bookings Report Excel file and produce a complete, professional HTML daily report email.

The report MUST include ALL of these sections with REAL numbers calculated from the data:

1. MTD SUMMARY: Total distinct orders (count unique Sales Order No values), total Net Sales USD (sum of Net Sales (USD) column), total Net Pounds (sum of Net weight where Net weight Unit = LB), monthly plan $25,000,000 (from $300M annual / 12), avg orders per day, avg net sales per day, date coverage range.

2. FORECAST VS PLAN: EOM run-rate forecast = (MTD sales / elapsed calendar days) * total days in month. Show variance vs $25M plan in dollars and percentage.

3. SHIPPING OUTLOOK: For each distinct Sales Order No, calculate lead days = EST Ship Date minus Created On Dt. Show average lead days, median lead days, average EST Ship Date, median EST Ship Date.

4. DAILY TABLE: Group orders by Created On Dt (date only). For each day show: distinct order count, net sales USD, net pounds LB, cumulative orders, cumulative sales, cumulative pounds.

5. TOP 5 CUSTOMERS: Group by Sold-To Name. Show customer name, net sales USD, order count, net pounds. Sort by net sales descending.

6. TOP 5 MATERIALS: Group by Material Text. Show material name, net sales USD, net pounds LB. Sort by net sales descending.

7. BUSINESS MIX: Group by PH1 column. Show PH1 code, net sales USD, order count. Sort by net sales descending. Show top 6.

8. SALESPERSON LEADERBOARD: Show a placeholder section saying the mapping file was not attached and listing the top customers by Sold-To Name as a proxy.

DESIGN REQUIREMENTS:
- Complete self-contained HTML with all CSS inline or in a <style> tag
- Dark navy header (#1F4E79) with white text
- Clean white body background
- KPI cards in a 3-column grid with colored left borders
- Tables with alternating row colors (#f8fafc for even rows)
- Professional fonts (Arial or system fonts only — no Google Fonts)
- Mobile-friendly
- All dollar amounts formatted with $ and commas (e.g. $1,234,567.89)
- All pound amounts formatted with commas and LB suffix

Return ONLY the complete HTML document. Start with <!DOCTYPE html>. No markdown, no explanation, nothing else."""
            },
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "data": esa_b64
                }
            }
        ]

        # Add mapping file if provided
        if map_file:
            map_b64 = base64.b64encode(map_file.read()).decode('utf-8')
            content[0]['text'] = content[0]['text'].replace(
                "8. SALESPERSON LEADERBOARD: Show a placeholder section saying the mapping file was not attached and listing the top customers by Sold-To Name as a proxy.",
                """8. SALESPERSON LEADERBOARD: Use the mapping file (second attachment) to join Sold-To No to Outside Salesperson, Inside Salesperson, and Sales Team columns.
Top 5 Outside Sales by net sales MTD. Top 5 Inside Sales by net sales MTD. Top 3 Sales Teams by net sales MTD.
Only include rows where Status = Active and salesperson is not blank or 'Need'."""
            )
            content.append({
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "data": map_b64
                }
            })

        # Call Claude
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            messages=[{"role": "user", "content": content}]
        )

        html = message.content[0].text
        # Strip markdown fences if present
        html = html.replace('```html', '').replace('```', '').strip()

        return jsonify({'html': html, 'success': True})

    except anthropic.AuthenticationError:
        return jsonify({'error': 'API key is invalid. Go to console.anthropic.com and create a new key.'}), 401
    except anthropic.RateLimitError:
        return jsonify({'error': 'Rate limit hit. Wait 30 seconds and try again.'}), 429
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
