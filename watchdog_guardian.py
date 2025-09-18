#!/usr/bin/env python3
"""
Watchdog Guardian - 監控並自動重啟 watchdog 主程式
當主程式崩潰時自動重新啟動，避免手動干預
"""
import subprocess
import time
import sys
import os
import datetime
import logging
from pathlib import Path
import threading

class WatchdogGuardian:
    def __init__(self, 
                 watchdog_script="main.py",
                 restart_delay=5,
                 max_restarts_per_hour=10,
                 log_file="guardian.log"):
        self.watchdog_script = watchdog_script
        self.restart_delay = restart_delay
        self.max_restarts_per_hour = max_restarts_per_hour
        self.log_file = log_file
        self.restart_times = []
        self.process = None
        self.dialog_killer_thread = None
        self.stop_dialog_killer = False
        
        # 設定 logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [Guardian] %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def kill_error_dialogs(self):
        """自動關閉 Windows 錯誤對話框（在背景線程中運行）"""
        if not sys.platform.startswith('win'):
            return
            
        try:
            import ctypes
            from ctypes import wintypes
            
            user32 = ctypes.windll.user32
            
            def enum_windows_proc(hwnd, lParam):
                # 獲取視窗標題
                length = user32.GetWindowTextLengthW(hwnd)
                if length == 0:
                    return True
                    
                buffer = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buffer, length + 1)
                window_title = buffer.value
                
                # 檢查是否為錯誤對話框
                error_titles = [
                    "python.exe", "Microsoft Visual C++", "Application Error",
                    "程式錯誤", "應用程式錯誤", "Runtime Error", "Fatal Error"
                ]
                
                if any(title.lower() in window_title.lower() for title in error_titles):
                    try:
                        # 嘗試關閉視窗
                        user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
                        self.logger.info(f"自動關閉錯誤對話框: {window_title}")
                    except Exception:
                        pass
                
                return True
            
            # 定義回調函數類型
            EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
            enum_proc = EnumWindowsProc(enum_windows_proc)
            
            while not self.stop_dialog_killer:
                try:
                    user32.EnumWindows(enum_proc, 0)
                except Exception:
                    pass
                time.sleep(2)  # 每2秒檢查一次
                
        except Exception as e:
            self.logger.warning(f"錯誤對話框殺手啟動失敗: {e}")

    def start_dialog_killer(self):
        """啟動錯誤對話框殺手線程"""
        if sys.platform.startswith('win') and not self.dialog_killer_thread:
            self.stop_dialog_killer = False
            self.dialog_killer_thread = threading.Thread(
                target=self.kill_error_dialogs, 
                daemon=True, 
                name="DialogKiller"
            )
            self.dialog_killer_thread.start()
            self.logger.info("錯誤對話框自動關閉功能已啟動")

    def stop_dialog_killer_func(self):
        """停止錯誤對話框殺手線程"""
        self.stop_dialog_killer = True
        if self.dialog_killer_thread:
            self.dialog_killer_thread = None

    def is_too_many_restarts(self):
        """檢查是否重啟次數過多"""
        now = datetime.datetime.now()
        one_hour_ago = now - datetime.timedelta(hours=1)
        
        # 清除一小時前的記錄
        self.restart_times = [t for t in self.restart_times if t > one_hour_ago]
        
        return len(self.restart_times) >= self.max_restarts_per_hour
    
    def start_watchdog(self):
        """啟動 watchdog 程式，跳過 UI 設定對話框"""
        try:
            # 使用 --auto-start 參數跳過 UI
            cmd = [sys.executable, self.watchdog_script, "--auto-start"]
            
            self.logger.info(f"啟動 watchdog: {' '.join(cmd)}")
            
            # Windows 特殊設定：禁用錯誤對話框
            startupinfo = None
            if sys.platform.startswith('win'):
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            # 使用 subprocess.Popen 以便監控程式狀態
            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1,
                startupinfo=startupinfo,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            )
            
            return True
            
        except Exception as e:
            self.logger.error(f"啟動 watchdog 失敗: {e}")
            return False
    
    def monitor_output(self):
        """監控程式輸出（可選）"""
        if not self.process:
            return
            
        try:
            # 非阻塞讀取一行輸出
            import select
            import fcntl
            import os
            
            # 設定非阻塞模式 (Unix/Linux)
            if hasattr(select, 'select'):
                fd = self.process.stdout.fileno()
                fl = fcntl.fcntl(fd, fcntl.F_GETFL)
                fcntl.fcntl(fd, fcntl.F_SETFL, fl | os.O_NONBLOCK)
                
                ready, _, _ = select.select([self.process.stdout], [], [], 0.1)
                if ready:
                    line = self.process.stdout.readline()
                    if line:
                        self.logger.info(f"[watchdog] {line.strip()}")
        except:
            pass  # Windows 或其他平台可能不支援
    
    def run(self):
        """主監控循環"""
        self.logger.info("Guardian 啟動，開始監控 watchdog")
        
        # 啟動錯誤對話框自動關閉功能
        self.start_dialog_killer()
        
        restart_count = 0
        
        while True:
            try:
                # 啟動 watchdog
                if not self.start_watchdog():
                    self.logger.error("無法啟動 watchdog，等待後重試")
                    time.sleep(self.restart_delay)
                    continue
                
                self.logger.info(f"Watchdog 已啟動 (PID: {self.process.pid})")
                
                # 監控程式運行狀態
                while True:
                    # 檢查程式是否仍在運行
                    returncode = self.process.poll()
                    
                    if returncode is not None:
                        # 程式已結束
                        self.logger.warning(f"Watchdog 程式結束，返回碼: {returncode}")
                        break
                    
                    # 監控輸出（可選）
                    self.monitor_output()
                    
                    # 短暫休眠
                    time.sleep(1)
                
                # 程式崩潰或結束，準備重啟
                restart_count += 1
                current_time = datetime.datetime.now()
                self.restart_times.append(current_time)
                
                # 檢查重啟頻率
                if self.is_too_many_restarts():
                    self.logger.error(f"一小時內重啟次數過多 ({len(self.restart_times)} 次)，暫停 1 小時")
                    time.sleep(3600)  # 等待一小時
                    self.restart_times.clear()
                    continue
                
                self.logger.info(f"準備重啟 watchdog (第 {restart_count} 次)，{self.restart_delay} 秒後重啟...")
                time.sleep(self.restart_delay)
                
            except KeyboardInterrupt:
                self.logger.info("收到停止信號，正在關閉...")
                self.stop_dialog_killer_func()
                if self.process and self.process.poll() is None:
                    self.logger.info("終止 watchdog 程式...")
                    self.process.terminate()
                    time.sleep(2)
                    if self.process.poll() is None:
                        self.process.kill()
                break
            except Exception as e:
                self.logger.error(f"Guardian 運行錯誤: {e}")
                time.sleep(self.restart_delay)

def main():
    print("=== Watchdog Guardian ===")
    print("監控並自動重啟 watchdog 程式")
    print("按 Ctrl+C 停止監控")
    print("=" * 30)
    
    guardian = WatchdogGuardian(
        watchdog_script="main.py",
        restart_delay=5,
        max_restarts_per_hour=10,
        log_file="guardian.log"
    )
    
    try:
        guardian.run()
    except KeyboardInterrupt:
        print("\nGuardian 已停止")

if __name__ == "__main__":
    main()