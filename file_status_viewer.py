#!/usr/bin/env python3
"""
檔案開啟/關閉狀態查詢介面 - 表格顯示
"""
import os
import sys
import json
import time
from datetime import datetime, timedelta
from typing import Dict, List, Optional

def format_duration(seconds: float) -> str:
    """格式化時間長度"""
    if seconds < 60:
        return f"{int(seconds)}秒"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        return f"{minutes}分{secs}秒"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        return f"{hours}時{minutes}分"

def print_table_separator(widths: List[int]):
    """印出表格分隔線"""
    parts = []
    for width in widths:
        parts.append("─" * width)
    print("├" + "┼".join(parts) + "┤")

def print_table_header(headers: List[str], widths: List[int]):
    """印出表格標題"""
    # 上邊框
    parts = []
    for width in widths:
        parts.append("─" * width)
    print("┌" + "┬".join(parts) + "┐")
    
    # 標題行
    row_parts = []
    for i, header in enumerate(headers):
        text = header[:widths[i]-2].center(widths[i]-2)
        row_parts.append(f" {text} ")
    print("│" + "│".join(row_parts) + "│")
    
    # 標題分隔線
    print_table_separator(widths)

def print_table_row(values: List[str], widths: List[int]):
    """印出表格行"""
    row_parts = []
    for i, value in enumerate(values):
        text = str(value)[:widths[i]-2].ljust(widths[i]-2)
        row_parts.append(f" {text} ")
    print("│" + "│".join(row_parts) + "│")

def print_table_footer(widths: List[int]):
    """印出表格底部"""
    parts = []
    for width in widths:
        parts.append("─" * width)
    print("└" + "┴".join(parts) + "┘")

def create_mock_data():
    """創建模擬數據用於展示"""
    current_time = datetime.now()
    return {
        r"C:\Users\user\Desktop\Test\A.xlsx": {
            'is_open': True,
            'temp_files': {'~$A.xlsx'},
            'opened_at': current_time - timedelta(minutes=15),
            'last_author': 'Michael Cheng'
        },
        r"C:\Users\user\Desktop\Test\B.xlsx": {
            'is_open': True,
            'temp_files': {'~$B.xlsx', 'B.tmp'},
            'opened_at': current_time - timedelta(minutes=8),
            'last_author': 'John Doe'
        },
        r"C:\Users\user\Desktop\Test\Data_v3.xlsx": {
            'is_open': False,
            'temp_files': set(),
            'opened_at': current_time - timedelta(hours=2),
            'last_author': 'Jane Smith'
        },
        r"C:\Users\user\Desktop\Test\Report.xlsx": {
            'is_open': False,
            'temp_files': set(),
            'opened_at': current_time - timedelta(hours=1, minutes=30),
            'last_author': 'Michael Cheng'
        }
    }

def show_open_files_table(file_status: Dict):
    """顯示開啟檔案的表格"""
    print("\n🟢 目前開啟的檔案")
    print("=" * 80)
    
    # 過濾開啟的檔案
    open_files = []
    current_time = datetime.now()
    
    for file_path, status in file_status.items():
        if status.get('is_open', False):
            filename = os.path.basename(file_path)
            author = status.get('last_author', '未知')
            opened_at = status.get('opened_at', current_time)
            duration = current_time - opened_at
            temp_count = len(status.get('temp_files', set()))
            
            open_files.append([
                filename,
                author,
                opened_at.strftime('%H:%M:%S'),
                format_duration(duration.total_seconds()),
                str(temp_count)
            ])
    
    if not open_files:
        print("   目前沒有檔案被開啟")
        return
    
    # 表格設定
    headers = ["檔案名稱", "使用者", "開啟時間", "開啟時長", "臨時檔"]
    widths = [25, 15, 12, 12, 8]
    
    # 印出表格
    print_table_header(headers, widths)
    
    for row in open_files:
        print_table_row(row, widths)
    
    print_table_footer(widths)
    
    print(f"\n📊 總計: {len(open_files)} 個檔案正在使用中")

def show_closed_files_table(file_status: Dict, limit: int = 10):
    """顯示最近關閉檔案的表格"""
    print(f"\n🔴 最近關閉的檔案 (最多 {limit} 個)")
    print("=" * 80)
    
    # 過濾關閉的檔案
    closed_files = []
    current_time = datetime.now()
    
    for file_path, status in file_status.items():
        if not status.get('is_open', True):  # 預設為 True，所以 False 表示關閉
            filename = os.path.basename(file_path)
            author = status.get('last_author', '未知')
            opened_at = status.get('opened_at', current_time)
            # 假設關閉時間為開啟時間後一段時間（實際需要記錄關閉時間）
            closed_at = opened_at + timedelta(minutes=20)  # 模擬數據
            duration = closed_at - opened_at
            
            closed_files.append([
                filename,
                author,
                opened_at.strftime('%H:%M:%S'),
                closed_at.strftime('%H:%M:%S'),
                format_duration(duration.total_seconds())
            ])
    
    if not closed_files:
        print("   沒有最近關閉的檔案記錄")
        return
    
    # 按開啟時間排序（最新在前）
    closed_files.sort(key=lambda x: x[2], reverse=True)
    closed_files = closed_files[:limit]
    
    # 表格設定
    headers = ["檔案名稱", "使用者", "開啟時間", "關閉時間", "使用時長"]
    widths = [25, 15, 12, 12, 12]
    
    # 印出表格
    print_table_header(headers, widths)
    
    for row in closed_files:
        print_table_row(row, widths)
    
    print_table_footer(widths)
    
    print(f"\n📊 顯示最近 {len(closed_files)} 個關閉的檔案")

def show_user_summary_table(file_status: Dict):
    """顯示用戶使用摘要表格"""
    print("\n👥 用戶使用摘要")
    print("=" * 60)
    
    # 統計用戶數據
    user_stats = {}
    
    for file_path, status in file_status.items():
        author = status.get('last_author', '未知')
        if author not in user_stats:
            user_stats[author] = {
                'open_files': 0,
                'total_files': 0
            }
        
        user_stats[author]['total_files'] += 1
        if status.get('is_open', False):
            user_stats[author]['open_files'] += 1
    
    if not user_stats:
        print("   沒有用戶數據")
        return
    
    # 準備表格數據
    user_rows = []
    for user, stats in user_stats.items():
        user_rows.append([
            user,
            str(stats['open_files']),
            str(stats['total_files']),
            f"{stats['open_files']/stats['total_files']*100:.1f}%" if stats['total_files'] > 0 else "0%"
        ])
    
    # 按開啟檔案數排序
    user_rows.sort(key=lambda x: int(x[1]), reverse=True)
    
    # 表格設定
    headers = ["使用者", "開啟中", "總檔案", "活躍率"]
    widths = [20, 10, 10, 10]
    
    # 印出表格
    print_table_header(headers, widths)
    
    for row in user_rows:
        print_table_row(row, widths)
    
    print_table_footer(widths)

def main():
    """主程式"""
    print("📊 Excel 檔案狀態監控表格")
    print("=" * 50)
    print("注意: 這是演示版本，使用模擬數據")
    print("=" * 50)
    
    # 使用模擬數據
    file_status = create_mock_data()
    
    # 顯示各種表格
    show_open_files_table(file_status)
    show_closed_files_table(file_status)
    show_user_summary_table(file_status)
    
    print("\n" + "=" * 50)
    print("💡 在實際使用中，這些數據來自 watchdog 程式的即時監控")
    print("💡 表格會顯示真實的檔案開啟/關閉狀態和用戶資訊")

if __name__ == "__main__":
    main()