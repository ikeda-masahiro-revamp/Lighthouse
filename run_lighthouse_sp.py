import subprocess
import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime


def run_lighthouse_for_url(url, output_dir):
    # 絶対パスを取得してディレクトリを作成
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    for i in range(1, 6):  # 5回計測
        output_path_json = os.path.abspath(os.path.join(output_dir, f"report_{i}.json"))
        log_path = os.path.abspath(os.path.join(output_dir, f"lighthouse_verbose_run_{i}.log"))

        print(f"Running Lighthouse for run {i} on URL: {url}")
        print(f"Output JSON path: {output_path_json}")
        print(f"Log path: {log_path}")

        command_json = [
            "npx", "lighthouse", url,
            "--preset=perf",
            "--output", "json",
            "--output-path", output_path_json,
            "--chrome-flags=--headless --no-sandbox",
            "--max-wait-for-load=60000",
            "--verbose",
            "--emulated-form-factor=mobile",
            "--throttling-method=simulate",
            "--throttling.cpuSlowdownMultiplier=2",
            "--throttling.throughputKbps=6000",
            "--throttling.uploadThroughputKbps=750",
            "--throttling.latency=100"
        ]

        with open(log_path, "w") as log_file:
            result_json = subprocess.run(command_json, stdout=log_file, stderr=subprocess.STDOUT)

        if result_json.returncode == 0:
            print(f"Lighthouse JSON run {i} completed successfully.")
        else:
            print(f"Error in Lighthouse JSON run {i}. Check {log_path} for details.")
            break  # エラー発生時は停止

    process_results(output_dir, url)


def process_results(output_dir, url):
    metrics_to_extract = [
        "FCP(ms)",
        "LCP(ms)",
        "CLS",
        "Speed Index(ms)",
        "TBT(ms)",
        "TTI(ms)",
        "TTFB(ms)",
        "Performance"
    ]

    results = []
    for i in range(1, 6):
        json_path = os.path.abspath(os.path.join(output_dir, f"report_{i}.json"))
        if os.path.exists(json_path):
            with open(json_path, 'r') as f:
                data = json.load(f)
                try:
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
                except KeyError as e:
                    print(f"Error extracting metrics from {json_path}: {e}")
        else:
            print(f"JSON file {json_path} not found!")

    if results:
        save_results_to_excel(output_dir, results, url)
    else:
        print("No metrics data found. Excel file not created.")


def save_results_to_excel(output_dir, results, url):
    df = pd.DataFrame(results)

    # 平均値を計算
    averages = pd.DataFrame([df.mean(numeric_only=True)]).assign(run="average")
    df = pd.concat([df, averages], ignore_index=True)

    # 実行完了日時を取得
    completed_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Excelファイルの書き出し
    excel_path = os.path.join(output_dir, f"{os.path.basename(output_dir)}.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name="Metrics")
        workbook = writer.book
        worksheet = writer.sheets["Metrics"]

        # URLと実行日時を追加
        worksheet["A1"] = f"URL: {url}"
        worksheet["A2"] = f"Completed at: {completed_datetime}"

        # BからG列の幅を設定
        for col_num in range(2, 9):  # B〜H
            col_letter = get_column_letter(col_num)
            worksheet.column_dimensions[col_letter].width = 15

        # 枠線を設定
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.row > 2:
                    cell.border = thin_border

        # 左揃えの設定（A列と1行目）
        for cell in worksheet["A"]:
            cell.alignment = Alignment(horizontal="left")

    print(f"計測結果のExcelファイルが次の場所に保存されました: {excel_path}")


if __name__ == "__main__":
    # 計測対象のURLリスト
    urls = [
        "https://furusato.jreast.co.jp/furusato",
        "https://furusato.jreast.co.jp/furusato/ranking"
    ]

    # URLごとに計測を実行
    for url in urls:
        output_dir = f"{url.split('//')[1].replace('/', '_')}"
        run_lighthouse_for_url(url, output_dir)
