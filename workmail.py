#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook 週報自動抓取彙整程式 (完整修正版)
功能：自動從 Outlook 抓取標題包含週報相關關鍵字的郵件並彙整
作者：Claude AI  
版本：1.5 (完整修正版)
"""

import win32com.client
import pandas as pd
from datetime import datetime, timedelta
import re
import os
import sys
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging

class OutlookReportExtractor:
    def setup_logging(self):
        """設定日誌記錄"""
        try:
            # 建立 logs 目錄（如果不存在）
            if not os.path.exists('logs'):
                os.makedirs('logs')
            
            # 設定日誌檔案
            log_file = os.path.join('logs', f'workmail_{datetime.now().strftime("%Y%m%d")}.log')
            
            # 設定日誌格式
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file, encoding='utf-8'),
                    logging.StreamHandler(sys.stdout)
                ]
            )
            self.logger = logging.getLogger(__name__)
        except Exception as e:
            print(f"設定日誌失敗: {str(e)}")
            self.logger = None

    def log_message(self, message, level='info'):
        """記錄訊息"""
        try:
            if self.logger:
                if level == 'error':
                    self.logger.error(message)
                else:
                    self.logger.info(message)
            print(message)
        except:
            pass  # 忽略記錄錯誤

    def __init__(self, account_email=None):
        """初始化 Outlook 連接"""
        self.setup_logging()
        
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.setup_account(account_email)
            self.log_message("成功連接到 Outlook")
            
        except Exception as e:
            error_msg = f"連接 Outlook 失敗: {e}"
            self.log_message(error_msg, 'error')
            self.log_message("請確認：")
            self.log_message("1. Microsoft Outlook 已開啟")
            self.log_message("2. 使用桌面版 Outlook（非網頁版）")
            self.log_message("3. 已成功登入帳戶")
            raise
    
    def setup_account(self, account_email):
        """設定要使用的帳戶"""
        try:
            accounts = self.outlook.Session.Accounts
            
            if account_email:
                self.target_account = None
                for account in accounts:
                    if account.SmtpAddress.lower() == account_email.lower():
                        self.target_account = account
                        break
                
                if not self.target_account:
                    raise Exception(f"找不到指定的帳戶: {account_email}")
                
                self.log_message(f"使用指定帳戶: {self.target_account.DisplayName} ({self.target_account.SmtpAddress})")
            else:
                if len(accounts) > 0:
                    self.target_account = accounts[0]
                    self.log_message(f"使用預設帳戶: {self.target_account.DisplayName} ({self.target_account.SmtpAddress})")
                else:
                    raise Exception("沒有找到任何 Outlook 帳戶")
                
                if len(accounts) > 1:
                    self.log_message("\n發現多個帳戶，目前使用第一個帳戶：")
                    for i, account in enumerate(accounts):
                        marker = "👉" if i == 0 else "  "
                        self.log_message(f"{marker} {account.DisplayName} ({account.SmtpAddress})")
                    self.log_message("如需使用其他帳戶，請在程式中指定 account_email 參數\n")
                    
        except Exception as e:
            raise Exception(f"設定帳戶失敗: {e}")
    
    def normalize_text(self, text):
        """正規化文字，移除空白、特殊符號並統一大小寫"""
        if not text:
            return ""
        
        # 移除空白字元、連字符、底線等
        normalized = re.sub(r'[\s\-_~・·．\.\,\，\(\)\[\]（）【】]', '', text)
        # 轉換為小寫
        normalized = normalized.lower()
        # 統一週/周字
        normalized = normalized.replace('週', '周')
        
        return normalized
    
    def is_weekly_report(self, subject):
        """判斷是否為週報郵件 - 優化版關鍵字比對"""
        if not subject:
            print(f"   ⏭️  跳過：主旨為空的郵件")
            return False
            
        # 優化後的關鍵字清單
        weekly_keywords = [
            # 直接關鍵字
            '周報', '週報', 'weekly', 'week',
            # 進度相關
            '工作進度', '进度', '進度報告', '進度匯報',
            '本周進度', '本週進度', '這周進度', '這週進度',
            '上周進度', '上週進度', '上周工作', '上週工作',
            '下周計畫', '下週計畫', '下周规划', '下週規劃',
            # 時間相關模式
            '周工作', '週工作', '周总结', '週總結',
            '工作总结', '工作總結', '工作匯報', '工作报告',
            # 英文相關
            'progress', 'report', 'summary', 'update'
        ]
        
        # 特殊模式匹配（正則表達式）
        special_patterns = [
            r'w\d+',           # w1, w2, w3 等
            r'week\d+',        # week1, week2 等
            r'\d+w\d+',        # 202506w1 等
            r'第\d+周',        # 第1周, 第2周 等
            r'第\d+週',        # 第1週, 第2週 等
            r'\d+/\d+.*進度',  # 6/2~6/6 工作進度 等
            r'\d+月\d+日.*進度', # 6月2日工作進度 等
        ]
        
        # 正規化主旨
        normalized_subject = self.normalize_text(subject)
        original_subject_lower = subject.lower()
        
        print(f"   🔍 原始主旨: {subject}")
        print(f"   🔧 正規化後: {normalized_subject}")
        
        # 檢查直接關鍵字匹配
        matched_keywords = []
        for keyword in weekly_keywords:
            normalized_keyword = self.normalize_text(keyword)
            
            # 在正規化文字中查找
            if normalized_keyword in normalized_subject:
                matched_keywords.append(keyword)
                print(f"   ✅ 符合關鍵字: '{keyword}' (正規化匹配)")
            # 在原始文字中查找（不區分大小寫）
            elif keyword.lower() in original_subject_lower:
                matched_keywords.append(keyword)
                print(f"   ✅ 符合關鍵字: '{keyword}' (原文匹配)")
        
        # 檢查特殊模式匹配
        matched_patterns = []
        for pattern in special_patterns:
            if re.search(pattern, normalized_subject, re.IGNORECASE):
                matched_patterns.append(pattern)
                print(f"   ✅ 符合模式: '{pattern}'")
            elif re.search(pattern, original_subject_lower, re.IGNORECASE):
                matched_patterns.append(pattern)
                print(f"   ✅ 符合模式: '{pattern}' (原文)")
        
        # 判斷結果
        is_report = len(matched_keywords) > 0 or len(matched_patterns) > 0
        
        if is_report:
            print(f"   🎯 識別為週報!")
            if matched_keywords:
                print(f"      📝 匹配關鍵字: {matched_keywords}")
            if matched_patterns:
                print(f"      🔍 匹配模式: {matched_patterns}")
        else:
            print(f"   ⏭️  跳過非週報信件")
            print(f"      ❌ 未符合任何關鍵字或模式")
        
        return is_report
    
    def get_emails_from_folder(self, folder_name="研發處", start_date=None, end_date=None):
        """從指定資料夾抓取郵件，遞迴搜尋所有子資料夾"""
        try:
            print(f"🔍 開始搜尋 Outlook 資料夾結構...")
            inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            print(f"📁 收件匣名稱: {inbox.Name}")
            
            # 列出收件匣中的所有資料夾
            print(f"\n📂 收件匣中的所有第一層資料夾:")
            try:
                for i, folder in enumerate(inbox.Folders, 1):
                    print(f"   {i}. 📁 '{folder.Name}'")
            except Exception as e:
                print(f"   ❌ 無法列出資料夾: {e}")

            # 直接從根目錄（與收件匣同層）尋找名稱為「研發處」的資料夾
            root_folder = inbox.Parent
            rd_folder = None
            for folder in root_folder.Folders:
                if folder.Name == folder_name:
                    rd_folder = folder
                    break

            if not rd_folder:
                print(f"❌ 未找到 '{folder_name}' 資料夾")
                return []

            print(f"✅ 成功找到目標資料夾: '{rd_folder.Name}'")

            # 遞迴收集所有子資料夾
            def collect_all_folders(folder):
                folders = [folder]
                try:
                    for subfolder in folder.Folders:
                        print(f"   📁 發現子資料夾: {subfolder.Name}")
                        folders.extend(collect_all_folders(subfolder))
                except Exception as e:
                    print(f"   ⚠️  讀取子資料夾時發生錯誤: {e}")
                return folders

            target_folders = collect_all_folders(rd_folder)
            print(f"📊 總共找到 {len(target_folders)} 個資料夾要搜尋")
            print("\n📂 搜尋的資料夾列表:")
            for i, folder in enumerate(target_folders, 1):
                print(f"   {i}. {folder.Name}")

            # 設定日期範圍
            if not start_date:
                end_date = datetime.now().replace(tzinfo=None)
                start_date = end_date - timedelta(days=end_date.weekday())  # 週一
                end_date = start_date + timedelta(days=6)  # 週日
            
            print(f"\n📅 搜尋時間範圍: {start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}")
            print(f"   📆 搜尋區間: {start_date.strftime('%Y/%m/%d')} 到 {end_date.strftime('%Y/%m/%d')}")

            all_messages = []
            total_emails_in_folders = 0
            folders_without_reports = []

            print(f"\n🔍 開始從各資料夾抓取郵件:")
            for i, folder in enumerate(target_folders, 1):
                try:
                    print(f"\n📁 [{i}/{len(target_folders)}] 處理資料夾: '{folder.Name}'")
                    
                    items = folder.Items
                    folder_total = len(items)
                    total_emails_in_folders += folder_total
                    print(f"   📊 資料夾總郵件數: {folder_total}")
                    
                    if folder_total == 0:
                        print(f"   ⚠️  資料夾是空的")
                        folders_without_reports.append(folder.Name)
                        continue
                    
                    # 統計時間範圍內的郵件
                    folder_messages = []
                    weekly_count = 0
                    
                    for j, item in enumerate(items):
                        try:
                            received_time = getattr(item, 'ReceivedTime', None)
                            if received_time:
                                received_time = received_time.replace(tzinfo=None)
                            subject = getattr(item, 'Subject', '') or ''
                            sender = getattr(item, 'SenderName', '') or ''
                            
                            # 檢查時間範圍
                            if received_time and start_date <= received_time <= end_date:
                                print(f"\n   📧 檢查信件 [{j+1}/{folder_total}]:")
                                print(f"   👤 寄件者: {sender}")
                                print(f"   📅 時間: {received_time.strftime('%Y/%m/%d %H:%M')}")
                                
                                # 使用優化後的週報判斷函數
                                if self.is_weekly_report(subject):
                                    weekly_count += 1
                                    print(f"   🎯 發現週報: [{sender}] {subject}")
                                    folder_messages.append(item)
                            
                            # 顯示前3封郵件的詳情用於調試
                            if j < 3:
                                time_str = received_time.strftime('%Y/%m/%d') if received_time else 'N/A'
                                in_range = "✅" if (received_time and start_date <= received_time <= end_date) else "❌"
                                print(f"   📋 [{j+1}] {in_range} {time_str} [{sender}] {subject[:50]}...")
                                
                        except Exception as e:
                            print(f"   ⚠️  讀取郵件 {j+1} 時發生錯誤: {e}")
                            continue
                    
                    print(f"   📧 時間範圍內郵件: {len(folder_messages)} 封")
                    print(f"   🎯 包含週報關鍵字: {weekly_count} 封")
                    
                    # 將符合條件的郵件加入總列表
                    all_messages.extend(folder_messages)

                    # 檢查資料夾是否有週報
                    if weekly_count == 0:
                        folders_without_reports.append(folder.Name)

                except Exception as e:
                    print(f"   ❌ 無法讀取資料夾 '{folder.Name}': {e}")
                    continue

            # 排序郵件
            all_messages.sort(key=lambda x: getattr(x, 'ReceivedTime', datetime.min), reverse=True)
            
            # 追蹤已提交和未提交週報的人員
            submitted_reports = set()
            all_senders = set()
            
            # 從所有資料夾中收集所有寄件者
            for folder in target_folders:
                try:
                    items = folder.Items
                    for item in items:
                        sender = getattr(item, 'SenderName', '') or ''
                        if sender:
                            all_senders.add(sender)
                except Exception as e:
                    print(f"   ⚠️  讀取寄件者資訊時發生錯誤: {e}")
            
            # 記錄已提交週報的人員
            for message in all_messages:
                sender = getattr(message, 'SenderName', '') or ''
                if sender:
                    submitted_reports.add(sender)
            
            # 計算未提交週報的人員
            missing_reports = all_senders - submitted_reports
            
            print(f"\n📊 搜尋結果統計:")
            print(f"   📁 搜尋資料夾數: {len(target_folders)}")
            print(f"   📧 資料夾總郵件數: {total_emails_in_folders}")
            print(f"   📅 時間範圍內郵件: {len(all_messages)}")
            
            # 顯示週報提交狀態
            print(f"\n📋 週報提交狀態:")
            if not submitted_reports:
                print("   ⚠️  尚未有人提交週報")
            else:
                print("   ✅ 已提交週報:")
                for sender in sorted(submitted_reports):
                    print(f"      • {sender}")
            
            if missing_reports:
                print("\n   ❌ 未提交週報:")
                for sender in sorted(missing_reports):
                    print(f"      • {sender}")
            
            # 列出未提供週報的資料夾
            if folders_without_reports:
                print(f"\n⚠️ 未提供週報的資料夾清單:")
                for folder in folders_without_reports:
                    print(f"   ❌ {folder}")
            
            self.logger.info(f"在 {len(target_folders)} 個資料夾中找到 {len(all_messages)} 封符合條件的週報")
            
            return all_messages

        except Exception as e:
            error_msg = f"抓取郵件失敗: {e}"
            print(f"❌ {error_msg}")
            self.logger.error(error_msg)
            return []
    
    def extract_report_content(self, message):
        """提取郵件內容並結構化"""
        try:
            def safe_get_property(obj, prop, default=""):
                try:
                    return getattr(obj, prop, default) or default
                except:
                    return default
            
            report_data = {
                'sender': safe_get_property(message, 'SenderName'),
                'sender_email': safe_get_property(message, 'SenderEmailAddress'),
                'subject': safe_get_property(message, 'Subject'),
                'received_time': safe_get_property(message, 'ReceivedTime'),
                'body': safe_get_property(message, 'Body'),
            }
            
            self.logger.info(f"成功提取報告: {report_data['sender']} - {report_data['subject'][:50]}")
            
            return report_data
            
        except Exception as e:
            error_msg = f"提取郵件內容失敗: {e}"
            self.logger.error(error_msg)
            print(f"⚠️  {error_msg}")
            return None
    
    def analyze_and_categorize_content(self, reports_data):
        """分析週報內容並分類整理 - 優化版避免重複和錯誤分類"""
        categorized_content = {
            'IScloud360_platform': [],
            'market_expansion': [],
            'esg_solutions': [],
            'cross_department': [],
            'technology_research': [],
            'other_projects': []
        }
        
        # 定義更精確的關鍵字分類規則（按優先級排序）
        category_keywords = {
            'esg_solutions': [
                # ESG相關 - 最高優先級，因為最具體
                'ESG', 'ESG-14064', 'ESG-14067', 'ESG-CBAM', 'ESG-GHG', 
                '14064', '14067', 'CBAM', 'GHG', '溫室氣體', '碳排放', '碳盤查',
                '永續', '環境', 'SaaS ESG', 'AP-SaaS', '合規體系'
            ],
            'cross_department': [
                # 跨部門支援 - 第二優先級（包含ERP和內部系統）
                '用友ERP', 'ERP導入', 'ERP系統', 'ERP建置', '內部系統', '系統導入',
                'Gotrust', '零信任產品', '零信任', 'CRM系統', 'EDMS系統', 
                '簽核流程', 'BUC報表', '出貨文件', '顧問文件', '跨部門支援', 
                '系統問題單', '流程優化', '系統建置', '內部流程', '業務報表'
            ],
            'IScloud360_platform': [
                # IScloud360相關 - 第三優先級（包含國內開發）
                'IScloud360', 'iscloud360', 'ISCloud', 'is cloud', 
                '平台升級', '平台優化', '行銷功能模組', '國際化布局', '三語系', '語系支援',
                '系統架構優化', '幣別參數', 'API串接採購', '自動化下單', '跨平台採購',
                '國內開發', '國內開發單', '平台開發', '功能開發', '模組開發'
            ],
            'technology_research': [
                # 前瞻技術相關 - 第四優先級
                'Copilot 4.0', 'Copilot', '代理人模式', 'AI CodeX', 'CodeX CLI',
                'Cursor AI', '開發工具評估', '前瞻技術', '能力提升', 'AI開發工具',
                '外訓課程', '技術深化', 'AI工具組合', '技術研究'
            ],
            'market_expansion': [
                # 市場拓展 - 最後優先級，關鍵字較廣泛
                '新創總會', 'AI雲市集', '供應商權限', '馬來西亞版本', '新加坡版本', 
                '日本版本', '海外版本', '國外市場', '市場拓展', '區域市場', '競爭優勢',
                '海外布局', '國際市場'
            ]
        }
        
        processed_reports = set()  # 用於追蹤已處理的報告，避免重複
        
        for report in reports_data:
            content = report.get('body', '')
            subject = report.get('subject', '')
            sender = report.get('sender', '')
            
            # 創建唯一標識符避免重複
            report_id = f"{sender}_{subject}_{len(content)}"
            if report_id in processed_reports:
                continue
            processed_reports.add(report_id)
            
            # 將主旨和內容合併分析
            full_text = f"{subject} {content}".lower()
            
            # 按優先級順序進行分類（找到第一個匹配就停止）
            categorized = False
            for category, keywords in category_keywords.items():
                if categorized:
                    break
                    
                for keyword in keywords:
                    if keyword.lower() in full_text:
                        categorized_content[category].append({
                            'sender': sender,
                            'subject': subject,
                            'content': content,
                            'matched_keyword': keyword,
                            'priority_score': self.calculate_priority_score(keyword, full_text)
                        })
                        categorized = True
                        print(f"   📂 分類: {category} - 關鍵字: {keyword} - {sender}")
                        break
            
            # 如果沒有匹配任何類別，放入其他專案
            if not categorized:
                categorized_content['other_projects'].append({
                    'sender': sender,
                    'subject': subject,
                    'content': content,
                    'matched_keyword': 'other',
                    'priority_score': 0
                })
                print(f"   📂 分類: other_projects - {sender}")
        
        # 對每個類別內的項目去重並排序
        for category in categorized_content:
            # 去除重複項目（基於sender和subject）
            unique_items = {}
            for item in categorized_content[category]:
                key = f"{item['sender']}_{item['subject']}"
                if key not in unique_items or item['priority_score'] > unique_items[key]['priority_score']:
                    unique_items[key] = item
            
            categorized_content[category] = list(unique_items.values())
            
            # 按優先級分數排序
            categorized_content[category].sort(key=lambda x: x['priority_score'], reverse=True)
        
        return categorized_content
    
    def calculate_priority_score(self, keyword, full_text):
        """計算關鍵字的優先級分數"""
        score = 0
        
        # 基礎分數
        score += full_text.count(keyword.lower()) * 10
        
        # 特定關鍵字加權
        high_priority_keywords = ['ESG', 'IScloud360', 'Copilot', 'ERP']
        if any(hpk.lower() in keyword.lower() for hpk in high_priority_keywords):
            score += 50
        
        # 在標題中出現給予更高分數
        if keyword.lower() in full_text.split('\n')[0].lower():
            score += 30
        
        return score
    
    def clean_content_text(self, text):
        """清理文字內容，移除不必要的前綴和格式"""
        if not text:
            return text
        
        # 常見的前綴模式
        prefix_patterns = [
            r'^[A-Z]+\s*-\s*',          # ISCOM - , ABC - 等
            r'^[A-Za-z0-9]+\s*:\s*',    # Project: , Task: 等
            r'^第\d+[項目點]\s*[\.、：:]\s*',  # 第1項. , 第2點：等
            r'^\d+\s*[\.、：:]\s*',      # 1. , 2：等
            r'^[•\*\-\+]\s*',           # • , * , - , + 等符號
            r'^\s*[\[\(].*?[\]\)]\s*',  # [標籤] , (分類) 等
            r'^Re:\s*',                 # Re: 回覆前綴
            r'^FW:\s*',                 # FW: 轉寄前綴
            r'^Fwd:\s*',                # Fwd: 轉寄前綴
        ]
        
        cleaned_text = text.strip()
        
        # 逐一移除匹配的前綴
        for pattern in prefix_patterns:
            cleaned_text = re.sub(pattern, '', cleaned_text, flags=re.IGNORECASE)
            cleaned_text = cleaned_text.strip()
        
        return cleaned_text
    
    def clean_final_text(self, text):
        """最終清理文字，移除多餘的標點和空白"""
        if not text:
            return text
        
        # 移除開頭的標點符號
        cleaned = re.sub(r'^[：:、，,\.\-\s]+', '', text)
        # 移除結尾的標點符號
        cleaned = re.sub(r'[：:、，,\.\-\s]+$', '', cleaned)
        # 移除重複的空白
        cleaned = re.sub(r'\s+', ' ', cleaned)
        
        return cleaned.strip()
    
    def extract_responsible_person(self, text, default_sender):
        """提取文字中的負責人資訊"""
        # 常見的負責人標示模式
        patterns = [
            r'\(([^)]+)\)$',           # 結尾括號內的名字 (jeffsun)
            r'負責人[：:]\s*([^\s,，]+)',  # 負責人：名字
            r'承辦人[：:]\s*([^\s,，]+)',  # 承辦人：名字
            r'執行人[：:]\s*([^\s,，]+)',  # 執行人：名字
            r'by\s+([^\s,，]+)',       # by 名字
            r'-\s*([^\s,，]+)$',       # 結尾 - 名字
        ]
        
        # 嘗試匹配各種模式
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                person = match.group(1).strip()
                # 過濾一些非人名的詞
                if not any(skip in person.lower() for skip in ['完成', '進行', '開發', '測試', '版本', '系統', '功能']):
                    return person
        
        # 如果沒有找到特定的負責人標示，返回郵件發送者
        return default_sender
    
    def extract_key_points(self, content_list):
        """從內容中提取關鍵要點並識別負責人 - 優化去重版"""
        key_points = []
        processed_contents = set()  # 用於去重
        
        for item in content_list:
            content = item['content']
            sender = item['sender']
            subject = item['subject']
            
            # 創建內容指紋避免重複處理相同內容
            content_fingerprint = f"{sender}_{subject}_{len(content)}"
            if content_fingerprint in processed_contents:
                continue
            processed_contents.add(content_fingerprint)
            
            # 簡單的關鍵點提取邏輯
            lines = content.split('\n')
            line_count = 0
            
            for line in lines:
                if line_count >= 3:  # 每個報告最多提取3個要點
                    break
                    
                line = line.strip()
                if line and len(line) > 15:  # 過濾太短的行
                    # 移除郵件簽名和無關內容
                    if any(skip in line.lower() for skip in [
                        '寄件者:', '收件者:', '主旨:', '-----', '此致', '敬祝', 
                        '最佳問候', 'best regards', '謝謝', '感謝', 'outlook'
                    ]):
                        continue
                    
                    # 清理內容 - 移除開頭的數字、符號
                    cleaned_line = re.sub(r'^[\d\.\-\*\•\s]+', '', line)
                    cleaned_line = cleaned_line.strip()
                    
                    # 使用清理函數移除前綴
                    cleaned_line = self.clean_content_text(cleaned_line)
                    
                    if len(cleaned_line) > 20:  # 保留有意義的內容
                        # 提取負責人資訊
                        responsible_person = self.extract_responsible_person(cleaned_line, sender)
                        
                        # 再次清理，移除負責人資訊後的重複內容
                        final_text = self.clean_final_text(cleaned_line)
                        
                        # 限制長度
                        if len(final_text) > 120:
                            final_text = final_text[:120] + "..."
                        
                        # 格式化要點，包含負責人
                        if final_text and len(final_text) > 10:  # 確保清理後還有有意義的內容
                            formatted_point = f"{final_text}({responsible_person})"
                            
                            # 檢查是否已存在相似的要點（避免重複）
                            is_duplicate = False
                            for existing_point in key_points:
                                if self.is_similar_content(final_text, existing_point.split('(')[0]):
                                    is_duplicate = True
                                    break
                            
                            if not is_duplicate:
                                key_points.append(formatted_point)
                                line_count += 1
        
        # 最終去重並限制數量
        unique_points = []
        for point in key_points:
            is_duplicate = False
            point_content = point.split('(')[0].strip()
            
            for existing in unique_points:
                existing_content = existing.split('(')[0].strip()
                if self.is_similar_content(point_content, existing_content):
                    is_duplicate = True
                    break
            
            if not is_duplicate:
                unique_points.append(point)
        
        return unique_points[:5]  # 最多5個要點
    
    def is_similar_content(self, text1, text2, threshold=0.7):
        """判斷兩個文字內容是否相似"""
        if not text1 or not text2:
            return False
        
        # 簡單的相似度計算
        text1_clean = re.sub(r'[^\w\s]', '', text1.lower())
        text2_clean = re.sub(r'[^\w\s]', '', text2.lower())
        
        # 如果其中一個是另一個的子集
        if text1_clean in text2_clean or text2_clean in text1_clean:
            return True
        
        # 計算共同詞彙比例
        words1 = set(text1_clean.split())
        words2 = set(text2_clean.split())
        
        if len(words1) == 0 or len(words2) == 0:
            return False
        
        common_words = words1.intersection(words2)
        similarity = len(common_words) / max(len(words1), len(words2))
        
        return similarity >= threshold
    
    def generate_summary_report(self, reports_data, output_file="週報彙整.docx", start_date=None, end_date=None, keywords=None):
        """產生彙整報告文件"""
        try:
            doc = Document()
            
            # 文件標題
            title = doc.add_heading('研發處週報彙整報告', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 產生資訊
            info_para = doc.add_paragraph()
            info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 使用介面設定的時間區間
            if start_date and end_date:
                period_text = f'報告期間: {start_date.strftime("%Y年%m月%d日")} ~ {end_date.strftime("%Y年%m月%d日")}'
            else:
                period_text = '報告期間: 無資料'
            
            info_para.add_run(f'產生時間: {datetime.now().strftime("%Y年%m月%d日 %H:%M")}').bold = True
            info_para.add_run(f'\n{period_text}')
            info_para.add_run(f'\n彙整報告數: {len(reports_data)} 份')
            
            # 人員統計區塊
            sender_stats = {}
            for report in reports_data:
                sender = report['sender']
                if sender in sender_stats:
                    sender_stats[sender] += 1
                else:
                    sender_stats[sender] = 1
            
            if sender_stats:
                doc.add_paragraph('')
                stat_title = doc.add_paragraph('人員統計:', style='Normal')
                stat_title.runs[0].bold = True
                stat_table = doc.add_table(rows=1, cols=2)
                stat_table.style = 'Table Grid'
                stat_table.cell(0, 0).text = '人員'
                stat_table.cell(0, 1).text = '週報數量'
                for sender, count in sorted(sender_stats.items()):
                    row_cells = stat_table.add_row().cells
                    row_cells[0].text = sender
                    row_cells[1].text = str(count)
            
            # 如果有設定關鍵字，顯示關鍵字資訊
            if keywords:
                keyword_para = doc.add_paragraph()
                keyword_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                keyword_para.add_run(f'\n關鍵字過濾: {", ".join(keywords)}').italic = True
            
            # 新增週報摘要
            doc.add_heading('週報摘要', level=1)
            
            # 依時間順序排列報告
            sorted_reports = sorted(reports_data, key=lambda x: x['received_time'], reverse=True)
            
            if keywords:
                doc.add_paragraph('（僅顯示包含關鍵字的內容）').italic = True
                
                # 在第一頁列出所有週報的摘要
                has_any_keyword_content = False
                for report in sorted_reports:
                    # 處理內容，將換行轉換為條列項目
                    content = report.get('body', '無內容')
                    if content:
                        # 移除多餘的空白行
                        lines = [line.strip() for line in content.split('\n') if line.strip()]
                        
                        # 建立條列清單
                        has_keyword_content = False
                        keyword_lines = []
                        
                        for line in lines:
                            # 清理內容
                            cleaned_line = self.clean_content_text(line)
                            if cleaned_line:
                                # 檢查是否包含關鍵字
                                if any(keyword.lower() in cleaned_line.lower() for keyword in keywords):
                                    has_keyword_content = True
                                    # 提取負責人資訊
                                    responsible_person = self.extract_responsible_person(cleaned_line, report["sender"])
                                    # 格式化為指定格式
                                    formatted_line = f"{cleaned_line}(ISCOM - {responsible_person})"
                                    keyword_lines.append(formatted_line)
                        
                        # 如果有符合關鍵字的內容，顯示報告資訊和內容
                        if has_keyword_content:
                            has_any_keyword_content = True
                            # 新增信件名片區塊
                            info_table = doc.add_table(rows=3, cols=2)
                            info_table.style = 'Table Grid'
                            
                            info_table.cell(0, 0).text = '報告人'
                            info_table.cell(0, 1).text = report["sender"]
                            info_table.cell(1, 0).text = '時間'
                            info_table.cell(1, 1).text = report["received_time"].strftime("%Y年%m月%d日 %H:%M")
                            info_table.cell(2, 0).text = '主旨'
                            info_table.cell(2, 1).text = report["subject"]
                            
                            # 顯示符合關鍵字的內容
                            for line in keyword_lines:
                                doc.add_paragraph(line, style='List Bullet')
                            
                            # 加入分隔線
                            doc.add_paragraph('─' * 60)
                
                # 如果沒有任何符合關鍵字的內容，顯示提示
                if not has_any_keyword_content:
                    doc.add_paragraph('（無任何符合關鍵字的內容）').italic = True
            else:
                doc.add_paragraph('（未設定關鍵字，不顯示摘要內容）').italic = True
            
            doc.add_page_break()
            
            # 依人員分組產生詳細報告
            doc.add_heading('週報詳細內容', level=1)
            
            for i, report in enumerate(sorted_reports, 1):
                # 個人週報標題
                doc.add_heading(f'{i}. {report["sender"]} - 週報', level=1)
                
                # 基本資訊
                info_table = doc.add_table(rows=3, cols=2)
                info_table.style = 'Table Grid'
                
                info_table.cell(0, 0).text = '報告人'
                info_table.cell(0, 1).text = report["sender"]
                info_table.cell(1, 0).text = '時間'
                info_table.cell(1, 1).text = report["received_time"].strftime("%Y年%m月%d日 %H:%M")
                info_table.cell(2, 0).text = '主旨'
                info_table.cell(2, 1).text = report["subject"]
                
                # 週報內容
                doc.add_heading('週報內容', level=2)
                content = report.get('body', '無內容')
                
                # 處理內容，將換行轉換為條列項目
                if content:
                    # 移除多餘的空白行
                    lines = [line.strip() for line in content.split('\n') if line.strip()]
                    
                    # 建立條列清單
                    for line in lines:
                        # 清理內容
                        cleaned_line = self.clean_content_text(line)
                        if cleaned_line:
                            # 提取負責人資訊
                            responsible_person = self.extract_responsible_person(cleaned_line, report["sender"])
                            # 格式化為指定格式
                            formatted_line = f"{cleaned_line}(ISCOM - {responsible_person})"
                            doc.add_paragraph(formatted_line, style='List Bullet')
                else:
                    doc.add_paragraph('無內容')
                
                # 分隔線
                if i < len(sorted_reports):
                    doc.add_paragraph('─' * 60)
                    doc.add_page_break()
            
            # 儲存文件
            doc.save(output_file)
            print(f"✅ Word 彙整報告已儲存至: {output_file}")
            self.logger.info(f"Word 報告已儲存: {output_file}")
            
        except Exception as e:
            error_msg = f"產生 Word 報告失敗: {e}"
            self.logger.error(error_msg)
            print(f"❌ {error_msg}")
    
    def run_extraction(self, folder_name="研發處", start_date=None, end_date=None, keywords=None):
        """執行完整的抓取和彙整流程"""
        print("🚀 開始週報抓取和彙整流程...")
        print("=" * 60)
        
        start_time = datetime.now()
        
        # 抓取郵件
        print("📬 正在抓取週報郵件...")
        messages = self.get_emails_from_folder(folder_name, start_date, end_date)
        
        if not messages or len(messages) == 0:
            print("❌ 沒有找到任何郵件")
            self.logger.warning("未找到任何郵件")
            return
        
        reports_data = []
        processed_count = 0
        skipped_count = 0
        
        print("\n🔍 正在分析郵件內容...")
        print("📋 週報識別標準: 只抓取標題包含「周報」或「週報」或「工作進度」或「本周進度」的信件")
        if keywords:
            print(f"🔍 關鍵字過濾: {', '.join(keywords)}")
        print("-" * 60)
        
        # 處理每封郵件
        for i, message in enumerate(messages, 1):
            try:
                subject = getattr(message, 'Subject', '') or ''
                sender = getattr(message, 'SenderName', '') or ''
                received_time = getattr(message, 'ReceivedTime', datetime.now())
                
                print(f"\n📧 檢查郵件 {i}/{len(messages)}: {sender}")
                print(f"   📅 時間: {received_time.strftime('%Y/%m/%d %H:%M') if hasattr(received_time, 'strftime') else str(received_time)}")
                print(f"   📝 主旨: {subject}")
                
                if self.is_weekly_report(subject):
                    print(f"   🎯 開始處理週報內容...")
                    
                report_data = self.extract_report_content(message)
                if report_data:
                    reports_data.append(report_data)
                    processed_count += 1
                    print(f"   ✅ 週報處理完成")
                else:
                    skipped_count += 1
                    print(f"   ❌ 週報處理失敗")
            except Exception as e:
                print(f"   ❌ 處理郵件時發生錯誤: {e}")
                skipped_count += 1
                continue
        
        if not reports_data:
            print("\n❌ 沒有找到標題包含「周報」或「週報」的信件")
            print("💡 建議檢查：")
            print("   - 郵件標題是否確實包含「周報」或「週報」字樣")
            print("   - 時間範圍是否正確")
            print("   - 郵件是否在研發處資料夾中")
            return
        
        print(f"\n📊 處理完成！")
        print(f"   ✅ 成功處理: {processed_count} 份週報")
        print(f"   ⏭️  跳過郵件: {skipped_count} 封")
        
        # 顯示處理的人員統計
        sender_stats = {}
        for report in reports_data:
            sender = report['sender']
            if sender in sender_stats:
                sender_stats[sender] += 1
            else:
                sender_stats[sender] = 1
        
        print(f"\n👥 人員統計:")
        for sender, count in sender_stats.items():
            print(f"   📝 {sender}: {count} 份週報")
        
        # 產生輸出檔案名稱
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        word_file = f"研發處_週報彙整_{timestamp}.docx"
        
        # 產生彙整報告
        print(f"\n📝 正在產生 Word 彙整報告...")
        self.generate_summary_report(reports_data, word_file, start_date, end_date, keywords)
        
        # 完成統計
        end_time = datetime.now()
        duration = end_time - start_time
        
        print("\n" + "=" * 60)
        print("🎉 週報彙整完成！")
        print(f"⏱️  總耗時: {duration.total_seconds():.1f} 秒")
        print(f"📁 輸出檔案:")
        print(f"   📄 Word 報告: {word_file}")
        
        self.logger.info(f"週報彙整完成，處理 {processed_count} 份報告，耗時 {duration.total_seconds():.1f} 秒")

def main():
    """主程式入口"""
    print("=" * 60)
    print("📧 Outlook 週報自動抓取彙整程式 (完整修正版)")
    print("=" * 60)
    
    try:
        # 程式設定
        folder_name = "研發處"
        start_date = None
        end_date = None
        
        print("⚙️  目前設定:")
        print(f"   📁 主資料夾: {folder_name}")
        print(f"   📅 搜尋範圍: 過去 {end_date - start_date if end_date else '全部'} 天")
        print(f"   🔍 識別規則: 智能識別週報相關關鍵字及格式模式")
        print(f"   📝 支援格式: 202506w1本周進度、6/2~6/6 Sheena工作進度、JoeWang_6W1工作進度 等")
        print(f"   🧹 內容清理: 自動移除 ISCOM- 等前綴，顯示負責人資訊")
        print()
        
        # 詢問是否要修改設定
        modify = input("是否要修改設定？(y/N): ").strip().lower()
        
        if modify in ['y', 'yes', '是']:
            new_folder = input(f"資料夾名稱 (目前: {folder_name}): ").strip()
            if new_folder:
                folder_name = new_folder
            
            try:
                start_input = input(f"開始日期 (目前: {start_date.strftime('%Y/%m/%d') if start_date else '全部'}): ").strip()
                if start_input:
                    start_date = datetime.strptime(start_input, '%Y/%m/%d').replace(tzinfo=None)
            except ValueError:
                print("⚠️  輸入的日期格式不正確，使用預設值")
            
            try:
                end_input = input(f"結束日期 (目前: {end_date.strftime('%Y/%m/%d') if end_date else '全部'}): ").strip()
                if end_input:
                    end_date = datetime.strptime(end_input, '%Y/%m/%d').replace(tzinfo=None)
            except ValueError:
                print("⚠️  輸入的日期格式不正確，使用預設值")
        
        print("\n🔧 使用設定:")
        print(f"   📁 主資料夾: {folder_name}")
        print(f"   📅 天數: {end_date - start_date if end_date else '全部'}")
        print()
        
        # 建立提取器實例
        print("🔗 正在連接 Outlook...")
        extractor = OutlookReportExtractor()
        
        # 執行抓取和彙整
        extractor.run_extraction(folder_name, start_date, end_date)
        
    except KeyboardInterrupt:
        print("\n\n⚠️  使用者中斷程式執行")
    except Exception as e:
        print(f"\n❌ 程式執行失敗: {e}")
        print("\n🔍 請檢查以下項目:")
        print("1. ✅ 已安裝必要套件: pip install pywin32 pandas python-docx")
        print("2. ✅ Microsoft Outlook 已開啟並正常運作")
        print("3. ✅ 使用桌面版 Outlook（非網頁版）")
        print("4. ✅ 已成功登入 Outlook 帳戶")
        print("5. ✅ 研發處資料夾確實存在")
        
        import traceback
        print(f"\n📋 詳細錯誤資訊:")
        print(traceback.format_exc())
        
    finally:
        input("\n按 Enter 鍵結束程式...")

if __name__ == "__main__":
    main()