import uiautomation as auto
import os
import time
import subprocess
import traceback
import configparser
import ctypes
import datetime

# pyinstaller -F autoVMware.py

def process_exists(process_name):
    '''
    Check if a program (based on its name) is running
    Return yes/no exists window and its PID
    '''
    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    # use buildin check_output right away
    try:
        output = subprocess.check_output(call, universal_newlines=True) # 在中文的console中使用需要解析編碼為big5???
        output = output.strip().split('\n')
        if len(output) == 1:  # 代表只有錯誤訊息
            return False, 0
        else:
            # check in last line for process name
            last_line_list = output[-1].lower().split()
        return last_line_list[0].startswith(process_name.lower()), int(last_line_list[1])
    except subprocess.CalledProcessError:
        return False, 0

config = configparser.ConfigParser()
config.read('autoVMware.ini', encoding='utf-8') 

ctypes.windll.kernel32.GetUserDefaultUILanguage()

ENGLISH_VERSION = True if (ctypes.windll.kernel32.GetUserDefaultUILanguage() == 1033) else False # 語言非英文目前視為中文
PROCESS_NAME = config['DEFAULT'].get('PROCESS_NAME')
PROCESS_PATH = config['DEFAULT'].get('PROCESS_PATH')
WINDOW_TARGET = config['DEFAULT'].get('WINDOW_TARGET')
ACCOUNT = config['DEFAULT'].get('ACCOUNT')
PASSWORD = config['DEFAULT'].get('PASSWORD')

auto.Logger.WriteLine(f"PROCESS_NAME:{PROCESS_NAME}\nPROCESS_PATH:{PROCESS_PATH}\nWINDOW_TARGET:{WINDOW_TARGET}", consoleColor=auto.ConsoleColor.Yellow)

while True:
    try:
        running, pid = process_exists(PROCESS_NAME)
        if running:
            window_target = auto.WindowControl(searchDepth=1, Name = WINDOW_TARGET, ProcessId=pid)
            if window_target.Exists(maxSearchSeconds=1): # 找到目標視窗
                auto.Logger.WriteLine(f"{datetime.datetime.today().strftime(r'%Y/%m/%d %H:%M:%S')}|TARGET WINDOW EXISTS\r", consoleColor=auto.ConsoleColor.Yellow, )
                # window_target.SetFocus() 會干擾其他程式使用
                bar = window_target.WindowControl(searchDepth=1, AutomationId = "ShadeBarWindow")
                if bar.Exists():
                    auto.Logger.WriteLine(f"ShadeBar EXISTS", consoleColor=auto.ConsoleColor.Yellow)
                    bar.SetFocus()
                    if ENGLISH_VERSION:
                        control = bar.CustomControl(searchDepth=1).ToolBarControl(searchDepth=1).ButtonControl(searchDepth=1, Name='Exit Fullscreen')
                    else:
                        control = bar.CustomControl(searchDepth=1).ToolBarControl(searchDepth=1).ButtonControl(searchDepth=1, Name='結束全螢幕')
                    if control.Exists():
                        control.GetInvokePattern().Invoke()
                        auto.Logger.WriteLine(f"Exit Fullscreen", consoleColor=auto.ConsoleColor.Yellow)
                time.sleep(60)
            else: # 沒找到目標視窗
                window = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client")
                if window.Exists():
                    control = window.CustomControl(searchDepth=1)
                    if control.Exists():
                        if control.ClassName == "WindowsPasswordAuthView": # 登入帳密輸入
                            if ENGLISH_VERSION:
                                control.EditControl(Name="Enter your user name").GetValuePattern().SetValue(ACCOUNT)
                                control.EditControl(Name="Enter your password").GetValuePattern().SetValue(PASSWORD)
                                control.ComboBoxControl(searchDepth=1).GetExpandCollapsePattern().Expand()
                                time.sleep(0.5)
                                control.ComboBoxControl(searchDepth=1).ListItemControl(searchDepth=1, Name = 'VGHTPE').GetSelectionItemPattern().Select()
                                control.ButtonControl(Name="Login").GetInvokePattern().Invoke()
                            else:
                                control.EditControl(Name="輸入您的使用者名稱").GetValuePattern().SetValue(ACCOUNT)
                                control.EditControl(Name="輸入您的密碼").GetValuePattern().SetValue(PASSWORD)
                                control.ComboBoxControl(searchDepth=1).GetExpandCollapsePattern().Expand()
                                time.sleep(0.5)
                                control.ComboBoxControl(searchDepth=1).ListItemControl(searchDepth=1, Name = 'VGHTPE').GetSelectionItemPattern().Select()
                                control.ButtonControl(Name="登入").GetInvokePattern().Invoke()
                        elif control.ClassName =='EntitlementsView': # 目前有連線狀況下開啟指定連線
                            list_control = control.ListControl(searchDepth=1)
                            list_item = list_control.ListItemControl(Name = WINDOW_TARGET)
                            if list_item.Exists():
                                list_item.SetFocus()
                                list_item.DoubleClick(simulateMove=False, waitTime=0.1)
                        elif control.ClassName == 'ServersView': # 目前無連線狀況下開啟指定連線
                            list_control = control.ListControl(searchDepth=1)
                            list_item = list_control.ListItemControl(searchDepth=1)
                            if list_item.Exists():
                                list_item.SetFocus()
                                list_item.DoubleClick(simulateMove=False, waitTime=0.1)
                            else:
                                if ENGLISH_VERSION:
                                    control.ToolBarControl(searchDepth=1).ButtonControl(searchDepth=1, Name='Add Server').GetInvokePattern().Invoke()
                                else:
                                    control.ToolBarControl(searchDepth=1).ButtonControl(searchDepth=1, Name='新增伺服器').GetInvokePattern().Invoke()
                        elif control.ClassName == 'NewServerView': # 新增伺服器
                            if ENGLISH_VERSION:
                                window = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client")
                                control.EditControl(searchDepth=1, Name='Name of the Connection Server').GetValuePattern().SetValue('vdt2.vghtpe.gov.tw')
                                control.ButtonControl(searchDepth=1, Name='Connect').GetInvokePattern().Invoke()
                            else:
                                window = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client")
                                control.EditControl(searchDepth=1, Name='連線伺服器的名稱').GetValuePattern().SetValue('vdt2.vghtpe.gov.tw')
                                control.ButtonControl(searchDepth=1, Name='連線').GetInvokePattern().Invoke()
                else:
                    w = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client Content Dialog")
                    if w.Exists():
                        w.GetWindowPattern().Close()
                    else:
                        auto.Logger.WriteLine(f"TARGET WINDOW NOT EXISTS, OPENING...", consoleColor=auto.ConsoleColor.Yellow)
                        os.startfile(PROCESS_PATH)
            time.sleep(1)
        else:
            auto.Logger.WriteLine(f"TARGET PROGRAM NOT EXISTS, OPENING...", consoleColor=auto.ConsoleColor.Yellow)
            os.startfile(PROCESS_PATH)
            time.sleep(1)
    except:
        time.sleep(1)
        auto.Logger.WriteLine(f"{traceback.print_exc()}", consoleColor=auto.ConsoleColor.Red)
        
