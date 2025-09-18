"""
看門狗監控腳本
獨立監控主程式的運行狀態，程式崩潰時自動重啟
"""
import os
import sys
import time
import json
import subprocess
import threading
from datetime import datetime, timedelta
import psutil


class WatchdogMonitor:
    """看門狗監控器"""
    
    def __init__(self):
        self.main_script = "main.py"
        self.heartbeat_file = "system_heartbeat.json"
        self.log_file = "watchdog_monitor.log"
        self.check_interval = 60  # 每分鐘檢查一次
        self.heartbeat_timeout = 120  # 2分鐘沒心跳就重啟
        self.restart_count = 0
        self.max_restarts_per_hour = 5
        self.restart_history = []
        
        self.running = True
        
    def log(self, message):
        """記錄日誌"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] {message}"
        print(log_entry)
        
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(log_entry + "\n")
        except Exception:
            pass
    
    def check_process_exists(self):
        """檢查主程式進程是否存在"""
        try:
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    cmdline = proc.info['cmdline']
                    if cmdline and len(cmdline) >= 2:
                        if 'python' in cmdline[0].lower() and self.main_script in ' '.join(cmdline):
                            return True, proc.info['pid']
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            return False, None
        except Exception as e:
            self.log(f"檢查進程時發生錯誤: {e}")
            return False, None
    
    def check_heartbeat(self):
        """檢查心跳檔案"""
        try:
            if not os.path.exists(self.heartbeat_file):
                return False, "心跳檔案不存在"
            
            with open(self.heartbeat_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            last_timestamp = data.get("timestamp", 0)
            current_time = time.time()
            elapsed = current_time - last_timestamp
            
            if elapsed > self.heartbeat_timeout:
                return False, f"心跳超時 {elapsed:.1f} 秒"
            
            return True, f"心跳正常，{elapsed:.1f} 秒前更新"
            
        except Exception as e:
            return False, f"讀取心跳檔案失敗: {e}"
    
    def check_restart_limit(self):
        """檢查重啟頻率限制"""
        current_time = datetime.now()
        one_hour_ago = current_time - timedelta(hours=1)
        
        # 清理一小時前的記錄
        self.restart_history = [t for t in self.restart_history if t > one_hour_ago]
        
        if len(self.restart_history) >= self.max_restarts_per_hour:
            return False, f"一小時內重啟次數過多 ({len(self.restart_history)}/{self.max_restarts_per_hour})"
        
        return True, "重啟頻率正常"
    
    def restart_main_program(self):
        """重啟主程式"""
        try:
            # 檢查重啟頻率
            can_restart, reason = self.check_restart_limit()
            if not can_restart:
                self.log(f"拒絕重啟: {reason}")
                return False
            
            self.log("正在重啟主程式...")
            
            # 嘗試優雅地終止現有進程
            process_exists, pid = self.check_process_exists()
            if process_exists:
                try:
                    proc = psutil.Process(pid)
                    proc.terminate()
                    proc.wait(timeout=10)
                    self.log(f"已終止現有進程 (PID: {pid})")
                except Exception as e:
                    self.log(f"終止現有進程失敗: {e}")
            
            # 啟動新進程
            startup_info = subprocess.STARTUPINFO()
            startup_info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startup_info.wShowWindow = subprocess.SW_NORMAL
            
            process = subprocess.Popen(
                [sys.executable, self.main_script],
                cwd=os.getcwd(),
                startupinfo=startup_info,
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            
            self.restart_count += 1
            self.restart_history.append(datetime.now())
            
            self.log(f"主程式已重啟 (PID: {process.pid}, 總重啟次數: {self.restart_count})")
            
            # 發送重啟通知
            self.send_restart_notification(True, process.pid)
            
            return True
            
        except Exception as e:
            self.log(f"重啟主程式失敗: {e}")
            self.send_restart_notification(False, None, str(e))
            return False
    
    def send_restart_notification(self, success, pid, error=None):
        """發送重啟通知"""
        try:
            from utils.email_notifier import send_notification
            
            if success:
                subject = "程式自動重啟成功"
                message = f"""
Excel 監控程式已自動重啟成功

重啟時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
新進程 PID: {pid}
總重啟次數: {self.restart_count}
最近一小時重啟: {len(self.restart_history)} 次

系統將繼續監控程式運行狀態。
                """
            else:
                subject = "程式自動重啟失敗"
                message = f"""
Excel 監控程式自動重啟失敗，需要人工干預

失敗時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
錯誤訊息: {error}
總重啟次數: {self.restart_count}

請檢查系統狀態並手動重啟程式。
                """
            
            send_notification(subject, message)
            
        except Exception as e:
            self.log(f"發送通知失敗: {e}")
    
    def monitor_loop(self):
        """主監控循環"""
        self.log("看門狗監控已啟動")
        self.log(f"監控目標: {self.main_script}")
        self.log(f"檢查間隔: {self.check_interval} 秒")
        self.log(f"心跳超時: {self.heartbeat_timeout} 秒")
        
        # 發送啟動通知
        try:
            from utils.email_notifier import send_notification
            send_notification(
                "看門狗監控已啟動",
                f"Excel 監控程式看門狗已於 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 啟動，開始監控主程式運行狀態。"
            )
        except Exception:
            pass
        
        while self.running:
            try:
                # 檢查進程是否存在
                process_exists, pid = self.check_process_exists()
                
                # 檢查心跳
                heartbeat_ok, heartbeat_msg = self.check_heartbeat()
                
                self.log(f"檢查結果 - 進程: {'存在' if process_exists else '不存在'} (PID: {pid}), 心跳: {heartbeat_msg}")
                
                # 判斷是否需要重啟
                need_restart = False
                restart_reason = ""
                
                if not process_exists:
                    need_restart = True
                    restart_reason = "主程式進程不存在"
                elif not heartbeat_ok:
                    need_restart = True
                    restart_reason = f"心跳異常: {heartbeat_msg}"
                
                if need_restart:
                    self.log(f"檢測到問題: {restart_reason}")
                    if self.restart_main_program():
                        self.log("重啟成功，等待程式穩定...")
                        time.sleep(30)  # 等待程式啟動穩定
                    else:
                        self.log("重啟失敗，等待下次檢查...")
                
                # 等待下次檢查
                time.sleep(self.check_interval)
                
            except KeyboardInterrupt:
                self.log("收到中斷信號，停止監控")
                self.running = False
                break
            except Exception as e:
                self.log(f"監控循環發生錯誤: {e}")
                time.sleep(self.check_interval)
    
    def stop(self):
        """停止監控"""
        self.running = False
        self.log("看門狗監控已停止")


def main():
    """主函數"""
    print("Excel 監控程式看門狗")
    print("=" * 50)
    
    monitor = WatchdogMonitor()
    
    try:
        monitor.monitor_loop()
    except KeyboardInterrupt:
        print("\n正在停止看門狗...")
    finally:
        monitor.stop()


if __name__ == "__main__":
    main()