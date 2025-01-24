import subprocess
import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# 環境変数からURLと出力ディレクトリを取得
url = os.getenv("URL", "https://furusato.jreast.co.jp/furusato/ranking/")  # デフォルトURL
output_dir = os.getenv("OUTPUT_DIR", "lighthouse_results")  # デフォルトの出力ディレクトリ
os.makedirs(output_dir, exist_ok=True)

# Lighthouse計測を5回実行
for i in range(1, 6):  # 5回計測
    output_path_json = os.path.join(output_dir, f"report_{i}.json")
    log_path = os.path.join(output_dir, f"lighthouse_verbose_run_{i}.log")  

    command_json = [
        "lighthouse", url,
        "--preset=perf",
        "--output=json",
        "--output-path", output_path_json,
        "--chrome-flags=--headless --no-sandbox",
        "--max-wait-for-load=60000",
        "--emulated-form-factor=mobile",
        "--throttling-method=simulate",
        "--throttling.cpuSlowdownMultiplier=2",
        "--throttling.throughputKbps=6000",
        "--throttling.uploadThroughputKbps=750",
        "--throttling.latency=100"
    ]

    print(f"Running Lighthouse for run {i} ...")
    with open(log_path, "w") as log_file:
        result_json = subprocess.run(command_json, stdout=log_file, stderr=subprocess.STDOUT)

    if result_json.returncode == 0:
        print(f"Lighthouse JSON run {i} completed successfully.")
    else:
        print(f"Error in Lighthouse JSON run {i}. Exiting. Check {log_path} for details.")
        break

# 計測結果から指標を抽出し、Excelに出力
metrics_to_extract = [
    "FCP(ms)", "LCP(ms)", "CLS", "Speed Index(ms)", "TBT(ms)", "TTI(ms)", "TTFB(ms)", "Performance"
]

results = []
for i in range(1, 6):
    json_path = os.path.join(output_dir, f"report_{i}.json")
    if os.path.exists(json_path):
        with open(json_path, 'r') as f:
            data = json.load(f)
            audits = data.get("audits", {})
            performance_score = data.get("categories", {}).get("performance", {}).get("score", "N/A")
            result = {
                "run": i,
                "Performance": round(performance_score * 100, 2) if isinstance(performance_score, (int, float)) else "N/A",
                "TTFB(ms)": audits.get("server-response-time", {}).get("numericValue", "N/A"),
                "FCP(ms)": audits.get("first-contentful-paint", {}).get("numericValue", "N/A"),
                "LCP(ms)": audits.get("largest-contentful-paint", {}).get("numericValue", "N/A"),
                "Speed Index(ms)": audits.get("speed-index", {}).get("numericValue", "N/A"),
                "TBT(ms)": audits.get("total-blocking-time", {}).get("numericValue", "N/A"),
                "TTI(ms)": audits.get("interactive", {}).get("numericValue", "N/A"),
                "CLS": audits.get("cumulative-layout-shift", {}).get("numericValue", "N/A")
            }
            results.append(result)
    else:
        print(f"JSON file {json_path} not found!")

# DataFrame作成とExcel出力
if results:
    df = pd.DataFrame(results)
    averages = pd.DataFrame([df.mean(numeric_only=True)]).assign(run="average")
    df = pd.concat([df, averages], ignore_index=True)
    completed_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    excel_path = os.path.join(output_dir, f"{output_dir}.xlsx")

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name="Metrics")
        workbook = writer.book
        worksheet = writer.sheets["Metrics"]
        worksheet["A1"] = f"URL: {url}"
        worksheet["A2"] = f"Completed at: {completed_datetime}"

        for col_num in range(2, 9):
            col_letter = get_column_letter(col_num)
            worksheet.column_dimensions[col_letter].width = 15

        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.row > 2:
                    cell.border = thin_border

    print(f"Excel report saved at {excel_path}")
else:
    print("No metrics data found. Excel file not created.")
