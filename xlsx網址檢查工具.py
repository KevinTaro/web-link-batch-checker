import pandas as pd
import requests
from openpyxl import load_workbook
import os

def check_url(url):
    try:
        resp = requests.get(url, timeout=10)
        return 'OK' if resp.status_code == 200 else f'HTTP {resp.status_code}'
    except requests.exceptions.RequestException as e:
        return str(e)

def main():
    xlsx_file = input('請輸入要檢查的 xlsx 檔案名稱：').strip()
    if not os.path.isabs(xlsx_file):
        xlsx_file = os.path.join('output', 'xlsx', xlsx_file)
    wb = load_workbook(xlsx_file)
    for ws in wb.worksheets:
        print(f'正在檢查工作表：{ws.title}')
        # 找到網址欄位
        header_row = 2 if ws.cell(row=2, column=1).value == '標題' else 1
        url_col = None
        for col in range(1, ws.max_column+1):
            if ws.cell(row=header_row, column=col).value == '網址':
                url_col = col
                break
        if not url_col:
            print(f'找不到「網址」欄位，略過工作表 {ws.title}')
            continue
        # 新增結果欄位
        result_col = ws.max_column + 1
        ws.cell(row=header_row, column=result_col, value='網址檢查結果')
        # 檢查每一列網址
        for row in range(header_row+1, ws.max_row+1):
            url = ws.cell(row=row, column=url_col).value
            if url:
                result = check_url(url)
                ws.cell(row=row, column=result_col, value=result)
    out_path = os.path.join('output', 'checked', f'檢查結果_{os.path.basename(xlsx_file)}')
    wb.save(out_path)
    print(f'已完成檢查，結果儲存為 {out_path}')

def run_project():
    print("=== 網頁批次抓取與網址檢查專案 ===")
    # 1. 執行網頁抓取
    import 網頁抓取工具
    網頁抓取工具.main()
    # 2. 執行網址檢查
    # 自動尋找 output 資料夾最新 xlsx
    xlsx_files = [f for f in os.listdir(os.path.join('output', 'xlsx')) if f.endswith('.xlsx') and f.startswith('打包網頁連結_')]
    if not xlsx_files:
        print('找不到要檢查的 xlsx 檔案')
        return
    xlsx_files.sort(reverse=True)
    xlsx_file = xlsx_files[0]
    print(f'自動選擇最新檔案：{xlsx_file}')
    wb = load_workbook(os.path.join('output', 'xlsx', xlsx_file))
    for ws in wb.worksheets:
        print(f'正在檢查工作表：{ws.title}')
        header_row = 2 if ws.cell(row=2, column=1).value == '標題' else 1
        url_col = None
        for col in range(1, ws.max_column+1):
            if ws.cell(row=header_row, column=col).value == '網址':
                url_col = col
                break
        if not url_col:
            print(f'找不到「網址」欄位，略過工作表 {ws.title}')
            continue
        result_col = ws.max_column + 1
        ws.cell(row=header_row, column=result_col, value='網址檢查結果')
        for row in range(header_row+1, ws.max_row+1):
            url = ws.cell(row=row, column=url_col).value
            if url:
                result = check_url(url)
                ws.cell(row=row, column=result_col, value=result)
    out_path = os.path.join('output', 'checked', f'檢查結果_{xlsx_file}')
    wb.save(out_path)
    print(f'已完成檢查，結果儲存為 {out_path}')

if __name__ == '__main__':
    run_project()
