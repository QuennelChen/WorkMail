#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 週報自動抓取彙整程式 (GUI版本)
功能：自動從 Outlook 抓取標題包含「週報」或「周報」的郵件並彙整
作者：Claude AI
版本：1.0 (GUI版本)
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import sys
from datetime import datetime, timedelta
from workmail import OutlookReportExtractor
import io
import logging

class RedirectText:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = io.StringIO()
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr
        self.queue = queue.Queue()
        self.running = True
        self.update_thread = threading.Thread(target=self._update_text)
        self.update_thread.daemon = True
        self.update_thread.start()

    def write(self, string):
        self.buffer.write(string)
        self.queue.put(string)
        self.old_stdout.write(string)

    def flush(self):
        self.buffer.flush()
        self.old_stdout.flush()

    def _update_text(self):
        while self.running:
            try:
                string = self.queue.get(timeout=0.1)
                self.text_widget.configure(state='normal')
                self.text_widget.insert(tk.END, string)
                self.text_widget.see(tk.END)
                self.text_widget.configure(state='disabled')
                self.text_widget.update_idletasks()
            except queue.Empty:
                continue

    def __enter__(self):
        sys.stdout = self
        sys.stderr = self
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.running = False
        if self.update_thread.is_alive():
            self.update_thread.join(timeout=1.0)
        sys.stdout = self.old_stdout
        sys.stderr = self.old_stderr

class WorkMailGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("研發處週報彙整工具")
        
        # 設定視窗最小尺寸
        self.root.minsize(800, 800)
        
        # 設定網格權重，使元件可以隨視窗調整大小
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # 建立主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.grid_rowconfigure(5, weight=1)  # 讓關鍵字列表可以垂直伸展
        main_frame.grid_columnconfigure(1, weight=1)  # 讓輸入框可以水平伸展
        
        # 資料夾名稱
        ttk.Label(main_frame, text="資料夾名稱:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.folder_name = ttk.Entry(main_frame, width=40)
        self.folder_name.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        self.folder_name.insert(0, "研發處")
        
        # 狀態列
        self.status_var = tk.StringVar()
        self.status_var.set("就緒")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=5, column=0, columnspan=2, sticky=tk.W)
        
        # 日期選擇區域
        date_frame = ttk.LabelFrame(main_frame, text="日期範圍", padding="5")
        date_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        date_frame.grid_columnconfigure(1, weight=1)
        
        # 開始日期
        ttk.Label(date_frame, text="開始日期:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.start_date = ttk.Entry(date_frame, width=20)
        self.start_date.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 結束日期
        ttk.Label(date_frame, text="結束日期:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.end_date = ttk.Entry(date_frame, width=20)
        self.end_date.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        # 設定預設為本週
        self.set_date_range("week")
        
        # 快速選擇按鈕
        quick_select_frame = ttk.Frame(date_frame)
        quick_select_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(quick_select_frame, text="本週", command=lambda: self.set_date_range("week")).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_select_frame, text="上週", command=lambda: self.set_date_range("last_week")).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_select_frame, text="本月", command=lambda: self.set_date_range("month")).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_select_frame, text="上月", command=lambda: self.set_date_range("last_month")).pack(side=tk.LEFT, padx=2)
        
        # 關鍵字設定區域
        keyword_frame = ttk.LabelFrame(main_frame, text="關注內容關鍵字", padding="5")
        keyword_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        keyword_frame.grid_columnconfigure(0, weight=1)
        keyword_frame.grid_rowconfigure(0, weight=1)
        
        # 關鍵字列表
        self.keyword_list = tk.Listbox(keyword_frame, height=6, selectmode=tk.EXTENDED)
        self.keyword_list.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 關鍵字輸入框和按鈕
        keyword_input_frame = ttk.Frame(keyword_frame)
        keyword_input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        keyword_input_frame.grid_columnconfigure(0, weight=1)
        
        self.keyword_entry = ttk.Entry(keyword_input_frame)
        self.keyword_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(keyword_input_frame, text="新增", command=self.add_keyword).grid(row=0, column=1, padx=2)
        ttk.Button(keyword_input_frame, text="刪除", command=self.remove_keyword).grid(row=0, column=2, padx=2)
        
        # 關鍵字操作按鈕
        keyword_action_frame = ttk.Frame(keyword_frame)
        keyword_action_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(keyword_action_frame, text="全選", command=self.select_all_keywords).pack(side=tk.LEFT, padx=2)
        ttk.Button(keyword_action_frame, text="取消全選", command=self.deselect_all_keywords).pack(side=tk.LEFT, padx=2)
        
        # 輸出區域
        output_frame = ttk.LabelFrame(main_frame, text="執行記錄", padding="5")
        output_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        output_frame.grid_columnconfigure(0, weight=1)
        output_frame.grid_rowconfigure(0, weight=1)
        
        # 輸出文字框
        self.output_text = tk.Text(output_frame, height=15, wrap=tk.WORD)
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.output_text.configure(state='disabled')
        
        # 捲軸
        scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=self.output_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.output_text.configure(yscrollcommand=scrollbar.set)
        
        # 執行按鈕
        ttk.Button(main_frame, text="執行", command=self.run_extraction).grid(row=4, column=0, columnspan=2, pady=10)
        
        # 預設關鍵字
        default_keywords = [
            "IScloud360", "ESG", "ERP", "系統", "開發", "測試",
            "功能", "模組", "API", "資料庫", "優化", "問題"
        ]
        for keyword in default_keywords:
            self.keyword_list.insert(tk.END, keyword)
        
        # 設定輸出重定向
        self.setup_logging()
        
        # 設定主視窗的初始尺寸
        self.root.geometry("1000x900")
        
        # 設定網格權重，使元件可以隨視窗調整大小
        main_frame.grid_rowconfigure(3, weight=3)
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)
    
    def set_date_range(self, range_type):
        """設定日期範圍"""
        today = datetime.now()
        
        if range_type == "week":
            # 本週（週一到週日）
            start = today - timedelta(days=today.weekday())
            end = start + timedelta(days=6)
            range_name = "本週"
        elif range_type == "last_week":
            # 上週（上週一到上週日）
            start = today - timedelta(days=today.weekday() + 7)
            end = start + timedelta(days=6)
            range_name = "上週"
        elif range_type == "month":
            # 本月（1號到月底）
            start = today.replace(day=1)
            if today.month == 12:
                end = today.replace(year=today.year + 1, month=1, day=1) - timedelta(days=1)
            else:
                end = today.replace(month=today.month + 1, day=1) - timedelta(days=1)
            range_name = "本月"
        elif range_type == "last_month":
            # 上月（上月1號到上月月底）
            if today.month == 1:
                start = today.replace(year=today.year - 1, month=12, day=1)
            else:
                start = today.replace(month=today.month - 1, day=1)
            end = today.replace(day=1) - timedelta(days=1)
            range_name = "上月"
        
        # 更新日期輸入框
        self.start_date.delete(0, tk.END)
        self.start_date.insert(0, start.strftime("%Y/%m/%d"))
        self.end_date.delete(0, tk.END)
        self.end_date.insert(0, end.strftime("%Y/%m/%d"))
        
        # 更新狀態列
        self.status_var.set(f"已設定{range_name}日期範圍：{start.strftime('%Y/%m/%d')} ~ {end.strftime('%Y/%m/%d')}")
    
    def setup_logging(self):
        """設定日誌輸出重定向"""
        self.redirect = RedirectText(self.output_text)
        self.redirect.__enter__()
    
    def add_keyword(self):
        keyword = self.keyword_entry.get().strip()
        if keyword and keyword not in self.keyword_list.get(0, tk.END):
            self.keyword_list.insert(tk.END, keyword)
            self.keyword_entry.delete(0, tk.END)
    
    def remove_keyword(self):
        """刪除選取的關鍵字"""
        selection = self.keyword_list.curselection()
        if selection:
            # 從後往前刪除，避免索引變化
            for index in sorted(selection, reverse=True):
                self.keyword_list.delete(index)
    
    def get_keywords(self):
        return list(self.keyword_list.get(0, tk.END))
    
    def run_extraction(self):
        try:
            folder_name = self.folder_name.get().strip()
            if not folder_name:
                messagebox.showerror("錯誤", "請輸入資料夾名稱")
                self.reset_ui()
                return
            
            # 解析日期
            try:
                start_date = datetime.strptime(self.start_date.get(), "%Y/%m/%d")
                end_date = datetime.strptime(self.end_date.get(), "%Y/%m/%d")
            except ValueError:
                messagebox.showerror("錯誤", "日期格式不正確，請使用 YYYY/MM/DD 格式")
                self.reset_ui()
                return
            
            if end_date < start_date:
                messagebox.showerror("錯誤", "結束日期不能早於開始日期")
                self.reset_ui()
                return
            
            # 更新狀態
            self.status_var.set("正在執行...")
            self.root.update()
            
            # 清空執行紀錄
            self.output_text.configure(state='normal')
            self.output_text.delete(1.0, tk.END)
            self.output_text.configure(state='disabled')
            
            # 執行彙整
            try:
                extractor = OutlookReportExtractor()
                # 傳遞時間區間參數和關鍵字
                extractor.run_extraction(
                    folder_name, 
                    start_date=start_date, 
                    end_date=end_date,
                    keywords=self.get_keywords()
                )
                
                # 完成後顯示訊息
                messagebox.showinfo("完成", "週報彙整完成！")
                
            except Exception as e:
                error_msg = f"執行時發生錯誤：{str(e)}"
                messagebox.showerror("錯誤", error_msg)
                logging.error(error_msg)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"程式執行錯誤：{str(e)}")
            logging.error(f"程式執行錯誤：{str(e)}")
        finally:
            self.reset_ui()
    
    def reset_ui(self):
        self.status_var.set("就緒")
        self.root.update()  # 確保UI更新
    
    def select_all_keywords(self):
        """全選關鍵字"""
        self.keyword_list.selection_set(0, tk.END)
    
    def deselect_all_keywords(self):
        """取消全選關鍵字"""
        self.keyword_list.selection_clear(0, tk.END)

    def write_log(self, message):
        """寫入執行紀錄"""
        self.output_text.configure(state='normal')
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.output_text.configure(state='disabled')
        self.root.update()  # 強制更新UI

def main():
    root = tk.Tk()
    app = WorkMailGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 