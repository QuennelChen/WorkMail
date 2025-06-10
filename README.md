# Outlook 週報自動抓取彙整工具

## 簡介
這是一個自動從 Microsoft Outlook 抓取週報相關郵件並彙整的工具。

## 功能特色
- 自動從 Outlook 抓取包含週報關鍵字的郵件
- 支援多種週報格式的解析
- 自動彙整並生成 Word 報告
- 提供圖形化使用者介面
- 支援日誌記錄功能

## 檔案結構
- `workmail.py` - 主程式核心邏輯
- `workmail_gui.py` - 圖形化使用者介面
- `build.py` - 建置腳本
- `一鍵打包.py` - 快速打包工具
- `requirements.txt` - Python 依賴套件列表
- `version.txt` - 版本資訊

## 環境需求
- Python 3.7+
- Microsoft Outlook (需安裝在本機)
- Windows 作業系統

## 安裝方式

1. 安裝 Python 依賴套件：
```bash
pip install -r requirements.txt
```

2. 確保 Microsoft Outlook 已安裝並設定完成

## 使用方式

### 命令列版本
```bash
python workmail.py
```

### 圖形化介面版本
```bash
python workmail_gui.py
```

### 一鍵打包
```bash
python 一鍵打包.py
```

## 建置執行檔
使用 PyInstaller 建置獨立執行檔：
```bash
python build.py
```

## 版本歷史
- v1.5 - 完整修正版，增加錯誤處理和日誌功能
- v1.0 - 初始版本

## 開發者
Claude AI

## 授權
本專案供內部使用。

## 注意事項
- 確保 Outlook 已正確設定郵件帳戶
- 首次使用時可能需要授權程式存取 Outlook
- 建議定期備份重要資料
