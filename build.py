import PyInstaller.__main__
import os
import sys

def build_exe():
    # 取得目前目錄
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 設定圖示路徑（如果有圖示的話）
    # icon_path = os.path.join(current_dir, 'icon.ico')
    
    # 設定打包參數
    params = [
        'workmail_gui.py',  # 主程式
        '--name=研發處週報彙整工具',  # 執行檔名稱
        '--onefile',  # 打包成單一執行檔
        '--noconsole',  # 不顯示命令列視窗
        '--clean',  # 清理暫存檔
        '--add-data=README.md;.',  # 加入說明文件
        # f'--icon={icon_path}',  # 加入圖示（如果有圖示的話）
        '--version-file=version.txt',  # 版本資訊
    ]
    
    # 執行打包
    PyInstaller.__main__.run(params)

if __name__ == '__main__':
    build_exe() 