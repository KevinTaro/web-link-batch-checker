import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import asyncio
import os
import queue


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
        self.msg_queue = queue.Queue()
        self.root.title('網頁批次抓取與網址檢查工具')
        self.root.geometry('700x600')

        self.url_label = tk.Label(root, text='請在下方輸入網址，每行一個：', fg='gray')
        self.url_label.pack(pady=(10,0))
        self.url_text = scrolledtext.ScrolledText(root, width=70, height=10)
        self.url_text.pack(pady=(0,10))

        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(pady=5)

        self.btn_grab = tk.Button(self.btn_frame, text='批次抓取網頁連結', command=self.start_grab)
        self.btn_grab.grid(row=0, column=0, padx=5)

        self.btn_check = tk.Button(self.btn_frame, text='檢查 Excel 網址', command=self.select_xlsx)
        self.btn_check.grid(row=0, column=1, padx=5)

        self.btn_all = tk.Button(self.btn_frame, text='一鍵執行完整流程', command=self.run_all)
        self.btn_all.grid(row=0, column=2, padx=5)


        self.btn_show_output = tk.Button(self.btn_frame, text='顯示輸出檔案位置', command=self.show_output_dir)
        self.btn_show_output.grid(row=0, column=4, padx=5)
        self.status = tk.Label(root, text='狀態：等待操作', fg='blue')
        self.status.pack(pady=5)

        self.result_label = tk.Label(root, text='執行結果與進度：', fg='gray')
        self.result_label.pack(pady=(10,0))
        self.result_text = scrolledtext.ScrolledText(root, width=70, height=15)
        self.result_text.pack(pady=(0,10))

    # self._cancel_flag = threading.Event()  # 取消功能已移除
        self.root.after(100, self.process_queue)

        self.last_output_file = None

    def process_queue(self):
        while not self.msg_queue.empty():
            msg, status = self.msg_queue.get()
            if msg:
                self.result_text.insert(tk.END, msg)
                self.result_text.see(tk.END)
            if status:
                self.status.config(text=status)
        current_status = self.status.cget('text')
        # 動畫只啟動一次
        if (current_status.startswith('狀態：正在檢查網址') or current_status.startswith('狀態：正在批次抓取')):
            if not hasattr(self, 'animating') or not self.animating:
                self.animating = True
                self._animate_status(0)
        else:
            self.animating = False
        self.root.after(100, self.process_queue)

    def animate_status_text(self, base_text, idx):
        dots = ['', '.', '..', '...']
        self.status.config(text=f'{base_text}{dots[idx % 4]}')
        self.root.update_idletasks()

    def _animate_status(self, idx):
        if getattr(self, 'animating', False):
            current_status = self.status.cget('text')
            if current_status.startswith('狀態：正在檢查網址'):
                self.animate_status_text('狀態：正在檢查網址', idx)
            elif current_status.startswith('狀態：正在批次抓取'):
                self.animate_status_text('狀態：正在批次抓取', idx)
            self.root.after(1000, lambda: self._animate_status(idx + 1))

    def get_latest_file(self, folder):
        if os.path.exists(folder):
            files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.xlsx')]
            if files:
                files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return files[0]
        return None

    # 取消功能已移除

    def show_output_dir(self):
        import subprocess
        if self.last_output_file and os.path.exists(self.last_output_file):
            subprocess.Popen(f'explorer /select,\"{self.last_output_file}\"')
        else:
            messagebox.showinfo('提示', '找不到本次產出的檔案！')

    def run_all(self):
        import time
        urls = self.url_text.get('1.0', tk.END).strip().splitlines()
        urls = [u for u in urls if u]
        if not urls:
            messagebox.showwarning('警告', '請輸入至少一個網址！')
            return
        self.msg_queue.put(('開始執行完整流程...\n', '狀態：正在執行完整流程...'))
        threading.Thread(target=self.all_worker, args=(urls,), daemon=True).start()

    def all_worker(self, urls):
        self._run_full_process(urls)

    def _run_full_process(self, urls):
        import time
        start_time = time.time()
        try:
            self.msg_queue.put(('正在批次抓取...\n', '狀態：正在批次抓取...'))
            web_grab_tool.gui_main(urls)
            self.msg_queue.put(('網頁連結已批次抓取並匯出至 output/csv 及 output/xlsx\n', '狀態：抓取完成，準備檢查網址...'))
            xlsx_dir = os.path.join('output', 'xlsx')
            xlsx_files = [f for f in os.listdir(xlsx_dir) if f.endswith('.xlsx') and f.startswith('打包網頁連結_')]
            if not xlsx_files:
                self.msg_queue.put(('找不到要檢查的 xlsx 檔案\n', '狀態：找不到要檢查的 xlsx 檔案'))
                return
            xlsx_files.sort(reverse=True)
            xlsx_file = os.path.join(xlsx_dir, xlsx_files[0])
            self.msg_queue.put(('正在檢查網址...\n', '狀態：正在檢查網址...'))
            xlsx_address_check_tool.gui_main(xlsx_file)
            checked_dir = os.path.join('output', 'checked')
            self.last_output_file = self.get_latest_file(checked_dir)
            elapsed = time.time() - start_time
            self.msg_queue.put((f'檢查結果已儲存於 output/checked\n', None))
            self.msg_queue.put((f'完整流程完成！總耗時：{elapsed:.1f} 秒\n', f'狀態：完整流程完成！總耗時 {elapsed:.1f} 秒，可再次輸入網址執行下一輪'))
            self.url_text.delete('1.0', tk.END)
        except Exception as e:
            elapsed = time.time() - start_time
            self.msg_queue.put((f'錯誤：{e}\n', None))
            self.msg_queue.put((f'完整流程失敗（耗時 {elapsed:.1f} 秒），可再次輸入網址執行下一輪\n', f'狀態：完整流程失敗（耗時 {elapsed:.1f} 秒），可再次輸入網址執行下一輪'))
            self.url_text.delete('1.0', tk.END)


    def start_grab(self):
        urls = self.url_text.get('1.0', tk.END).strip().splitlines()
        urls = [u for u in urls if u]
        if not urls:
            messagebox.showwarning('警告', '請輸入至少一個網址！')
            return
        self.msg_queue.put(('正在批次抓取...\n', '狀態：正在批次抓取...'))
        threading.Thread(target=self.grab_worker, args=(urls,), daemon=True).start()

    def grab_worker(self, urls):
        self._run_grab_only(urls)

    def _run_grab_only(self, urls):
        try:
            self.msg_queue.put(('正在批次抓取...\n', '狀態：正在批次抓取...'))
            web_grab_tool.gui_main(urls)
            xlsx_dir = os.path.join('output', 'xlsx')
            self.last_output_file = self.get_latest_file(xlsx_dir)
            self.msg_queue.put(('網頁連結已批次抓取並匯出至 output/csv 及 output/xlsx\n', '狀態：抓取完成！'))
        except Exception as e:
            self.msg_queue.put((f'錯誤：{e}\n', '狀態：抓取失敗'))

    def select_xlsx(self):
        xlsx_path = filedialog.askopenfilename(title='選擇要檢查的 Excel 檔案', filetypes=[('Excel Files', '*.xlsx')])
        if not xlsx_path:
            return
        self.msg_queue.put(('正在檢查網址...\n', '狀態：正在檢查網址...'))
        threading.Thread(target=self.check_worker, args=(xlsx_path,), daemon=True).start()

    def check_worker(self, xlsx_path):
        self._run_check_only(xlsx_path)

    def _run_check_only(self, xlsx_path):
        try:
            self.msg_queue.put(('正在檢查網址...\n', '狀態：正在檢查網址...'))
            xlsx_address_check_tool.gui_main(xlsx_path)
            checked_dir = os.path.join('output', 'checked')
            self.last_output_file = self.get_latest_file(checked_dir)
            self.msg_queue.put((f'檢查結果已儲存於 output/checked\n', '狀態：檢查完成！'))
        except Exception as e:
            self.msg_queue.put((f'錯誤：{e}\n', '狀態：檢查失敗'))

if __name__ == '__main__':
    root = tk.Tk()
    app = WebLinkCheckerGUI(root)
    root.mainloop()
