py -m pip install --upgrade pip
py -m pip install pyinstaller pywin32 pandas python-docx
cd C:\Users\RexChang\source\testTool\Python開發工具\workmail
py -m PyInstaller --onefile --noconsole --name "週報彙整工具" --hidden-import win32timezone --hidden-import win32com.client --hidden-import pandas --hidden-import docx --hidden-import logging --hidden-import datetime --hidden-import re --hidden-import os --hidden-import sys --hidden-import win32com.client.gencache --hidden-import win32com.client.makepy --hidden-import win32com.client.util --hidden-import win32com.client.dynamic workmail.py