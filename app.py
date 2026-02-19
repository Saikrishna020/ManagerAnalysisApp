from flask import Flask, render_template_string, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Manager Case Analysis</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #6366f1;
            --primary-dark: #4f46e5;
            --primary-light: #818cf8;
            --accent: #06b6d4;
            --success: #10b981;
            --success-dark: #059669;
            --danger: #ef4444;
            --bg: #f8fafc;
            --surface: #ffffff;
            --text: #1e293b;
            --text-secondary: #64748b;
            --border: #e2e8f0;
            --shadow: 0 1px 3px rgba(0,0,0,0.06), 0 6px 16px rgba(0,0,0,0.04);
            --shadow-lg: 0 4px 6px rgba(0,0,0,0.04), 0 10px 40px rgba(0,0,0,0.08);
            --radius: 16px;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background: var(--bg);
            color: var(--text);
            min-height: 100vh;
        }

        /* Animated gradient header */
        .header {
            background: linear-gradient(135deg, #6366f1 0%, #06b6d4 50%, #8b5cf6 100%);
            background-size: 200% 200%;
            animation: gradientShift 8s ease infinite;
            padding: 48px 20px 42px;
            text-align: center;
            position: relative;
            overflow: hidden;
        }
        .header::before {
            content: '';
            position: absolute;
            top: -50%;
            left: -50%;
            width: 200%;
            height: 200%;
            background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 60%);
            animation: float 6s ease-in-out infinite;
        }
        @keyframes gradientShift {
            0%, 100% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
        }
        @keyframes float {
            0%, 100% { transform: translate(0, 0) rotate(0deg); }
            50% { transform: translate(30px, -20px) rotate(5deg); }
        }
        .header h1 {
            font-size: 32px;
            font-weight: 700;
            color: white;
            position: relative;
            z-index: 1;
            letter-spacing: -0.5px;
        }
        .header p {
            color: rgba(255,255,255,0.85);
            font-size: 15px;
            margin-top: 8px;
            position: relative;
            z-index: 1;
            font-weight: 400;
        }

        .container { max-width: 900px; margin: -30px auto 40px; padding: 0 20px; position: relative; z-index: 2; }

        /* Upload card */
        .card {
            background: var(--surface);
            border-radius: var(--radius);
            padding: 32px;
            box-shadow: var(--shadow);
            margin-bottom: 24px;
            border: 1px solid var(--border);
            transition: box-shadow 0.3s ease;
        }
        .card:hover { box-shadow: var(--shadow-lg); }

        .form-group { margin-bottom: 24px; }
        .form-group label {
            font-weight: 600;
            font-size: 14px;
            color: var(--text);
            display: block;
            margin-bottom: 8px;
            letter-spacing: 0.01em;
        }
        .form-group .hint {
            font-size: 12px;
            color: var(--text-secondary);
            margin-bottom: 8px;
        }

        /* Custom file upload */
        .file-upload-wrapper {
            position: relative;
            border: 2px dashed var(--border);
            border-radius: 12px;
            padding: 32px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background: #fafbff;
        }
        .file-upload-wrapper:hover {
            border-color: var(--primary-light);
            background: #f0f0ff;
        }
        .file-upload-wrapper.dragover {
            border-color: var(--primary);
            background: #ede9fe;
        }
        .file-upload-wrapper input[type="file"] {
            position: absolute;
            top: 0; left: 0; width: 100%; height: 100%;
            opacity: 0;
            cursor: pointer;
        }
        .file-upload-icon { font-size: 36px; margin-bottom: 8px; }
        .file-upload-text { font-size: 14px; color: var(--text-secondary); }
        .file-upload-text strong { color: var(--primary); }
        .file-name {
            margin-top: 10px;
            font-size: 13px;
            color: var(--success);
            font-weight: 500;
            display: none;
        }

        /* Select dropdown */
        .custom-select {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid var(--border);
            border-radius: 12px;
            font-size: 15px;
            font-family: inherit;
            color: var(--text);
            background: white;
            appearance: none;
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='20' height='20' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");
            background-repeat: no-repeat;
            background-position: right 14px center;
            cursor: pointer;
            transition: border-color 0.2s;
        }
        .custom-select:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(99,102,241,0.1); }

        /* Buttons */
        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 14px 32px;
            border: none;
            border-radius: 12px;
            font-size: 15px;
            font-weight: 600;
            font-family: inherit;
            cursor: pointer;
            transition: all 0.2s ease;
            text-decoration: none;
        }
        .btn-primary {
            background: linear-gradient(135deg, var(--primary), var(--primary-dark));
            color: white;
            box-shadow: 0 4px 12px rgba(99,102,241,0.3);
        }
        .btn-primary:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 20px rgba(99,102,241,0.4);
        }
        .btn-primary:active { transform: translateY(0); }
        .btn-success {
            background: linear-gradient(135deg, var(--success), var(--success-dark));
            color: white;
            box-shadow: 0 4px 12px rgba(16,185,129,0.3);
        }
        .btn-success:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 20px rgba(16,185,129,0.4);
        }
        .btn-block { width: 100%; justify-content: center; }

        /* Stats ribbon */
        .stats-ribbon {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 16px;
            margin-bottom: 24px;
        }
        .stat-card {
            background: var(--surface);
            border-radius: var(--radius);
            padding: 20px;
            text-align: center;
            border: 1px solid var(--border);
            box-shadow: var(--shadow);
        }
        .stat-value {
            font-size: 28px;
            font-weight: 700;
            background: linear-gradient(135deg, var(--primary), var(--accent));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        .stat-label { font-size: 12px; color: var(--text-secondary); margin-top: 4px; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 500; }

        /* Tables */
        .table-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 16px;
        }
        .table-title {
            font-size: 18px;
            font-weight: 700;
            color: var(--text);
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .table-title .icon {
            width: 32px;
            height: 32px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
        }
        .icon-blue { background: #ede9fe; }
        .icon-teal { background: #ccfbf1; }
        .badge {
            font-size: 12px;
            font-weight: 600;
            padding: 4px 12px;
            border-radius: 20px;
            background: #f1f5f9;
            color: var(--text-secondary);
        }
        table { width: 100%; border-collapse: separate; border-spacing: 0; }
        thead th {
            background: #f8fafc;
            padding: 12px 20px;
            text-align: left;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-secondary);
            border-bottom: 2px solid var(--border);
        }
        thead th:first-child { border-radius: 10px 0 0 0; }
        thead th:last-child { border-radius: 0 10px 0 0; text-align: right; }
        tbody td {
            padding: 14px 20px;
            font-size: 14px;
            border-bottom: 1px solid #f1f5f9;
        }
        tbody td:last-child { text-align: right; font-weight: 600; font-variant-numeric: tabular-nums; }
        tbody tr { transition: background 0.15s; }
        tbody tr:hover { background: #f8fafc; }
        tbody tr:last-child td { border-bottom: none; }

        /* rank badges for top 3 */
        .rank { display: inline-flex; align-items: center; gap: 8px; }
        .rank-badge {
            width: 24px; height: 24px; border-radius: 50%; display: inline-flex;
            align-items: center; justify-content: center; font-size: 11px; font-weight: 700; color: white;
        }
        .rank-1 { background: linear-gradient(135deg, #f59e0b, #d97706); }
        .rank-2 { background: linear-gradient(135deg, #94a3b8, #64748b); }
        .rank-3 { background: linear-gradient(135deg, #c2884d, #92603a); }

        /* Case count bar */
        .case-count {
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 10px;
        }
        .count-bar {
            height: 6px;
            border-radius: 3px;
            background: linear-gradient(90deg, var(--primary-light), var(--accent));
            min-width: 4px;
            transition: width 0.6s ease;
        }

        /* Download section */
        .download-section {
            text-align: center;
            padding: 28px;
        }
        .download-section p {
            font-size: 13px;
            color: var(--text-secondary);
            margin-top: 12px;
        }

        /* Error */
        .error-card {
            background: #fef2f2;
            border: 1px solid #fecaca;
            border-radius: var(--radius);
            padding: 20px 24px;
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 24px;
        }
        .error-icon { font-size: 22px; }
        .error-text { color: #991b1b; font-size: 14px; font-weight: 500; }

        /* Footer */
        .footer {
            text-align: center;
            padding: 24px;
            color: var(--text-secondary);
            font-size: 13px;
        }
        .footer a { color: var(--primary); text-decoration: none; }

        /* Animations */
        @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .animate { animation: fadeInUp 0.5s ease forwards; }
        .delay-1 { animation-delay: 0.1s; opacity: 0; }
        .delay-2 { animation-delay: 0.2s; opacity: 0; }
        .delay-3 { animation-delay: 0.3s; opacity: 0; }
        .delay-4 { animation-delay: 0.4s; opacity: 0; }

        /* Responsive */
        @media (max-width: 640px) {
            .header { padding: 36px 16px 32px; }
            .header h1 { font-size: 24px; }
            .container { margin-top: -20px; padding: 0 12px; }
            .card { padding: 20px; }
            .stats-ribbon { grid-template-columns: 1fr; }
            .btn { padding: 12px 24px; font-size: 14px; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Manager Case Analysis</h1>
        <p>Upload your Excel file and get instant manager workload insights</p>
    </div>

    <div class="container">
        <!-- Upload Form -->
        <div class="card animate delay-1">
            <form method="POST" enctype="multipart/form-data" action="/analyze" id="uploadForm">
                <div class="form-group">
                    <label for="file">Upload Excel File</label>
                    <div class="hint">Accepted format: .xlsx</div>
                    <div class="file-upload-wrapper" id="dropZone">
                        <input type="file" name="file" id="file" accept=".xlsx" required>
                        <div class="file-upload-icon">📁</div>
                        <div class="file-upload-text"><strong>Click to upload</strong> or drag and drop</div>
                        <div class="file-name" id="fileName"></div>
                    </div>
                </div>

                <div class="form-group">
                    <label for="analysis_type">Analysis Type</label>
                    <select name="analysis_type" id="analysis_type" class="custom-select">
                        <option value="Report Manager" {% if report_col == 'Report Manager' %}selected{% endif %}>📋 Report Manager</option>
                        <option value="Assigning Manager" {% if report_col == 'Assigning Manager' %}selected{% endif %}>👤 Assigning Manager</option>
                        <option value="Allotment Manager" {% if report_col == 'Allotment Manager' %}selected{% endif %}>📌 Allotment Manager</option>
                    </select>
                </div>

                <button type="submit" class="btn btn-primary btn-block">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
                    Analyze
                </button>
            </form>
        </div>

        {% if error %}
        <div class="error-card animate">
            <div class="error-icon">⚠️</div>
            <div class="error-text">{{ error }}</div>
        </div>
        {% endif %}

        {% if manager_cases is not none %}
        <!-- Stats -->
        <div class="stats-ribbon animate delay-1">
            <div class="stat-card">
                <div class="stat-value">{{ manager_cases|length }}</div>
                <div class="stat-label">Managers</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{{ total_cases }}</div>
                <div class="stat-label">Total Cases</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{{ other_cases|length }}</div>
                <div class="stat-label">{{ report_col }}s</div>
            </div>
        </div>

        <!-- Manager Cases Table -->
        <div class="card animate delay-2">
            <div class="table-header">
                <div class="table-title">
                    <span class="icon icon-blue">👔</span> Manager Cases
                </div>
                <span class="badge">{{ manager_cases|length }} managers</span>
            </div>
            <table>
                <thead><tr><th>Manager</th><th>Cases</th></tr></thead>
                <tbody>
                {% for row in manager_cases %}
                    <tr>
                        <td>
                            <span class="rank">
                                {% if loop.index <= 3 %}<span class="rank-badge rank-{{ loop.index }}">{{ loop.index }}</span>{% endif %}
                                {{ row[0] }}
                            </span>
                        </td>
                        <td>
                            <div class="case-count">
                                <div class="count-bar" style="width: {{ (row[1] / max_manager * 100)|int }}px;"></div>
                                {{ row[1] }}
                            </div>
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>

        <!-- Report/Assigning/Allotment Manager Cases Table -->
        <div class="card animate delay-3">
            <div class="table-header">
                <div class="table-title">
                    <span class="icon icon-teal">📊</span> {{ report_col }} Cases
                </div>
                <span class="badge">{{ other_cases|length }} managers</span>
            </div>
            <table>
                <thead><tr><th>{{ report_col }}</th><th>Cases</th></tr></thead>
                <tbody>
                {% for row in other_cases %}
                    <tr>
                        <td>
                            <span class="rank">
                                {% if loop.index <= 3 %}<span class="rank-badge rank-{{ loop.index }}">{{ loop.index }}</span>{% endif %}
                                {{ row[0] }}
                            </span>
                        </td>
                        <td>
                            <div class="case-count">
                                <div class="count-bar" style="width: {{ (row[1] / max_other * 100)|int }}px;"></div>
                                {{ row[1] }}
                            </div>
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>

        <!-- Download -->
        <div class="card download-section animate delay-4">
            <form method="POST" action="/download">
                <input type="hidden" name="manager_data" value="{{ manager_data }}">
                <input type="hidden" name="other_data" value="{{ other_data }}">
                <input type="hidden" name="report_col" value="{{ report_col }}">
                <input type="hidden" name="output_file" value="{{ output_file }}">
                <button type="submit" class="btn btn-success">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                    Download Analysis Report
                </button>
                <p>Excel file with all analysis sheets</p>
            </form>
        </div>
        {% endif %}

        <div class="footer">Made with care &middot; Manager Case Analysis Tool</div>
    </div>

    <script>
        // File upload interaction
        const fileInput = document.getElementById('file');
        const fileName = document.getElementById('fileName');
        const dropZone = document.getElementById('dropZone');

        fileInput.addEventListener('change', function() {
            if (this.files.length) {
                fileName.textContent = '✅ ' + this.files[0].name;
                fileName.style.display = 'block';
                dropZone.style.borderColor = '#10b981';
                dropZone.style.background = '#f0fdf4';
            }
        });

        dropZone.addEventListener('dragover', function(e) {
            e.preventDefault();
            this.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', function() {
            this.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', function() {
            this.classList.remove('dragover');
        });
    </script>
</body>
</html>
'''

@app.route("/")
def home():
    return render_template_string(HTML_TEMPLATE, manager_cases=None, error=None)

@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        file = request.files.get("file")
        analysis_type = request.form.get("analysis_type")

        if not file:
            return render_template_string(HTML_TEMPLATE, manager_cases=None, error="Please upload an Excel file.")

        df = pd.read_excel(file, skiprows=1)
        df.columns = df.columns.str.strip()

        if analysis_type == "Report Manager":
            df["Report Manager"] = df["Report Manager"].fillna(df["Manager"])
            report_col = "Report Manager"
            output_file = "manager_case_analysis_report.xlsx"
        elif analysis_type == "Assigning Manager":
            df["Assigning Manager"] = df["Assigning Manager"].fillna(df["Manager"])
            report_col = "Assigning Manager"
            output_file = "manager_case_analysis_assigning.xlsx"
        else:
            df["Allotment Manager"] = df["Allotment Manager"].fillna(df["Manager"])
            report_col = "Allotment Manager"
            output_file = "manager_case_analysis_allotment.xlsx"

        manager_cases = df["Manager"].value_counts().reset_index()
        manager_cases.columns = ["Manager", "Number of Cases"]

        other_cases = df[report_col].value_counts().reset_index()
        other_cases.columns = [report_col, "Number of Cases"]

        # Convert to lists for template
        manager_list = manager_cases.values.tolist()
        other_list = other_cases.values.tolist()

        # Encode data for download
        import json
        manager_data = json.dumps(manager_list)
        other_data = json.dumps(other_list)

        return render_template_string(
            HTML_TEMPLATE,
            manager_cases=manager_list,
            other_cases=other_list,
            report_col=report_col,
            output_file=output_file,
            manager_data=manager_data,
            other_data=other_data,
            total_cases=int(sum(r[1] for r in manager_list)),
            max_manager=max(r[1] for r in manager_list) if manager_list else 1,
            max_other=max(r[1] for r in other_list) if other_list else 1,
            error=None
        )
    except Exception as e:
        return render_template_string(HTML_TEMPLATE, manager_cases=None, error=str(e))

@app.route("/download", methods=["POST"])
def download():
    import json
    try:
        manager_data = json.loads(request.form.get("manager_data"))
        other_data = json.loads(request.form.get("other_data"))
        report_col = request.form.get("report_col")
        output_file = request.form.get("output_file")

        manager_cases = pd.DataFrame(manager_data, columns=["Manager", "Number of Cases"])
        other_cases = pd.DataFrame(other_data, columns=[report_col, "Number of Cases"])

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            manager_cases.to_excel(writer, sheet_name="Manager Cases", index=False)
            other_cases.to_excel(writer, sheet_name=f"{report_col} Cases", index=False)

        buffer.seek(0)
        return send_file(
            buffer,
            as_attachment=True,
            download_name=output_file,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return f"Error generating download: {e}", 500

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
