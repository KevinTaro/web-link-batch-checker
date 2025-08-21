import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import asyncio
import os

# 假設已經有 async 版本的批次檢查與抓取
import xlsx_address_check_tool
import web_grab_tool

def run_async(func, *args, **kwargs):
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(func(*args, **kwargs))
    loop.close()

class WebLinkCheckerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('網頁批次抓取與網址檢查工具')
        self.root.geometry('600x500')

        self.url_text = scrolledtext.ScrolledText(root, width=70, height=10)
        self.url_text.pack(pady=10)
        self.url_text.insert(tk.END, '請在此輸入網址，每行一個...')

        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(pady=5)

        self.btn_grab = tk.Button(self.btn_frame, text='批次抓取網頁連結', command=self.start_grab)
        self.btn_grab.grid(row=0, column=0, padx=5)

        self.btn_check = tk.Button(self.btn_frame, text='檢查 Excel 網址', command=self.select_xlsx)
        self.btn_check.grid(row=0, column=1, padx=5)

        self.btn_all = tk.Button(self.btn_frame, text='一鍵執行完整流程', command=self.run_all)
        self.btn_all.grid(row=0, column=2, padx=5)
        self.status = tk.Label(root, text='', fg='blue')
        self.status.pack(pady=5)

        self.result_text = scrolledtext.ScrolledText(root, width=70, height=15)
        self.result_text.pack(pady=10)

    def run_all(self):
        urls = self.url_text.get('1.0', tk.END).strip().splitlines()
        urls = [u for u in urls if u and not u.startswith('請在此輸入')]
        if not urls:
            messagebox.showwarning('警告', '請輸入至少一個網址！')
            return
        self.status.config(text='正在執行完整流程...')
        threading.Thread(target=self.all_worker, args=(urls,), daemon=True).start()

    def all_worker(self, urls):
        try:
            # 1. 抓取
            self.status.config(text='正在批次抓取...')
            web_grab_tool.gui_main(urls)
            self.result_text.insert(tk.END, '網頁連結已批次抓取並匯出至 output/csv 及 output/xlsx\n')
            # 2. 自動選擇最新 xlsx 檔案
            xlsx_dir = os.path.join('output', 'xlsx')
            xlsx_files = [f for f in os.listdir(xlsx_dir) if f.endswith('.xlsx') and f.startswith('打包網頁連結_')]
            if not xlsx_files:
                self.status.config(text='找不到要檢查的 xlsx 檔案')
                return
            xlsx_files.sort(reverse=True)
            xlsx_file = os.path.join(xlsx_dir, xlsx_files[0])
            self.status.config(text='正在檢查網址...')
            xlsx_address_check_tool.gui_main(xlsx_file)
            self.status.config(text='完整流程完成！')
            self.result_text.insert(tk.END, f'檢查結果已儲存於 output/checked\n')
        except Exception as e:
            self.status.config(text='完整流程失敗')
            self.result_text.insert(tk.END, f'錯誤：{e}\n')

        self.status = tk.Label(root, text='', fg='blue')
        self.status.pack(pady=5)

        self.result_text = scrolledtext.ScrolledText(root, width=70, height=15)
        self.result_text.pack(pady=10)

    def start_grab(self):
        urls = self.url_text.get('1.0', tk.END).strip().splitlines()
        urls = [u for u in urls if u and not u.startswith('請在此輸入')]
        if not urls:
            messagebox.showwarning('警告', '請輸入至少一個網址！')
            return
        self.status.config(text='正在批次抓取...')
        threading.Thread(target=self.grab_worker, args=(urls,), daemon=True).start()

    def grab_worker(self, urls):
        try:
            # 執行 async GUI 入口
            web_grab_tool.gui_main(urls)
            self.status.config(text='抓取完成！')
            self.result_text.insert(tk.END, '網頁連結已批次抓取並匯出至 output/csv 及 output/xlsx\n')
        except Exception as e:
            self.status.config(text='抓取失敗')
            self.result_text.insert(tk.END, f'錯誤：{e}\n')

    def select_xlsx(self):
        xlsx_path = filedialog.askopenfilename(title='選擇要檢查的 Excel 檔案', filetypes=[('Excel Files', '*.xlsx')])
        if not xlsx_path:
            return
        self.status.config(text='正在檢查網址...')
        threading.Thread(target=self.check_worker, args=(xlsx_path,), daemon=True).start()

    def check_worker(self, xlsx_path):
        try:
            # 執行 async GUI 入口
            xlsx_address_check_tool.gui_main(xlsx_path)
            self.status.config(text='檢查完成！')
            self.result_text.insert(tk.END, f'檢查結果已儲存於 output/checked\n')
        except Exception as e:
            self.status.config(text='檢查失敗')
            self.result_text.insert(tk.END, f'錯誤：{e}\n')

if __name__ == '__main__':
    root = tk.Tk()
    app = WebLinkCheckerGUI(root)
    root.mainloop()
