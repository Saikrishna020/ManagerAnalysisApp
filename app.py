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
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f6; color: #262730; }
        .container { max-width: 800px; margin: 40px auto; padding: 20px; }
        h1 { text-align: center; margin-bottom: 10px; color: #262730; }
        .subtitle { text-align: center; color: #666; margin-bottom: 30px; }
        .card { background: white; border-radius: 10px; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.08); margin-bottom: 20px; }
        label { font-weight: 600; display: block; margin-bottom: 8px; font-size: 15px; }
        select, input[type="file"] { width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 14px; margin-bottom: 20px; }
        select:focus, input[type="file"]:focus { outline: none; border-color: #4F8BF9; }
        .btn { background: #4F8BF9; color: white; border: none; padding: 12px 28px; border-radius: 6px; font-size: 15px; cursor: pointer; display: inline-block; text-decoration: none; }
        .btn:hover { background: #3a7be0; }
        .btn-download { background: #28a745; margin-top: 15px; }
        .btn-download:hover { background: #218838; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th { background: #4F8BF9; color: white; padding: 10px 15px; text-align: left; font-size: 14px; }
        td { padding: 10px 15px; border-bottom: 1px solid #eee; font-size: 14px; }
        tr:hover { background: #f8f9fa; }
        .section-title { font-size: 18px; font-weight: 600; margin: 25px 0 10px 0; color: #333; }
        .error { background: #ffe0e0; color: #c00; padding: 15px; border-radius: 6px; margin-top: 15px; }
        .footer { text-align: center; margin-top: 30px; color: #999; font-size: 13px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Manager Case Analysis Tool</h1>
        <p class="subtitle">Upload your Excel file and select the type of manager analysis you'd like to perform.</p>

        <div class="card">
            <form method="POST" enctype="multipart/form-data" action="/analyze">
                <label for="file">Upload Excel File</label>
                <input type="file" name="file" id="file" accept=".xlsx" required>

                <label for="analysis_type">Select Analysis Type</label>
                <select name="analysis_type" id="analysis_type">
                    <option value="Report Manager">Report Manager</option>
                    <option value="Assigning Manager">Assigning Manager</option>
                    <option value="Allotment Manager">Allotment Manager</option>
                </select>

                <button type="submit" class="btn">🔍 Analyze</button>
            </form>
        </div>

        {% if error %}
        <div class="error">⚠️ {{ error }}</div>
        {% endif %}

        {% if manager_cases is not none %}
        <div class="card">
            <div class="section-title">🔹 Manager Cases</div>
            <table>
                <thead><tr><th>Manager</th><th>Number of Cases</th></tr></thead>
                <tbody>
                {% for row in manager_cases %}
                    <tr><td>{{ row[0] }}</td><td>{{ row[1] }}</td></tr>
                {% endfor %}
                </tbody>
            </table>
        </div>

        <div class="card">
            <div class="section-title">🔹 {{ report_col }} Cases</div>
            <table>
                <thead><tr><th>{{ report_col }}</th><th>Number of Cases</th></tr></thead>
                <tbody>
                {% for row in other_cases %}
                    <tr><td>{{ row[0] }}</td><td>{{ row[1] }}</td></tr>
                {% endfor %}
                </tbody>
            </table>
        </div>

        <div class="card" style="text-align:center;">
            <form method="POST" action="/download">
                <input type="hidden" name="manager_data" value="{{ manager_data }}">
                <input type="hidden" name="other_data" value="{{ other_data }}">
                <input type="hidden" name="report_col" value="{{ report_col }}">
                <input type="hidden" name="output_file" value="{{ output_file }}">
                <button type="submit" class="btn btn-download">📥 Download Analysis Report</button>
            </form>
        </div>
        {% endif %}

        <p class="footer">Manager Case Analysis Tool</p>
    </div>
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
