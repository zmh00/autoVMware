import uiautomation as auto
import os
import time
import subprocess
import traceback
import configparser
import ctypes
import datetime
import shutil
from pathlib import Path
import sys

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


def main():
    global FOUND, NOT_FOUND, RESIZE_ALREADY
    while True:
        time.sleep(SLEEP_BEFORE_FOUND)
        try:
            running, pid = process_exists(PROCESS_NAME)
            if running:
                window_target = auto.WindowControl(searchDepth=1, Name = WINDOW_TARGET, ProcessId=pid)
                if window_target.Exists(maxSearchSeconds=1): # 找到目標視窗
                    if FOUND == 0:
                        auto.Logger.WriteLine(f"{datetime.datetime.today()}|TARGET WINDOW FOUND!", consoleColor=auto.ConsoleColor.Yellow)
                        FOUND = 1
                    
                    NOT_FOUND = 0
                    # auto.Logger.WriteLine(f"{datetime.datetime.today().strftime(r'%Y/%m/%d %H:%M:%S')}|TARGET WINDOW EXISTS\r", consoleColor=auto.ConsoleColor.Yellow, )
                    # window_target.SetFocus() 會干擾其他程式使用
                    
                    if not RESIZE_ALREADY:
                        # 關閉全螢幕模式
                        bar = window_target.WindowControl(searchDepth=1, AutomationId = "ShadeBarWindow")
                        if bar.Exists():
                            auto.Logger.WriteLine(f"{datetime.datetime.today()}|ShadeBar EXISTS", consoleColor=auto.ConsoleColor.Yellow)
                            bar.SetFocus()
                            if ENGLISH_VERSION:
                                control = bar.CustomControl(searchDepth=1).ToolBarControl(searchDepth=1).ButtonControl(searchDepth=1, Name='Exit Fullscreen')
                            else:
                                control = bar.CustomControl(searchDepth=1).ToolBarControl(searchDepth=1).ButtonControl(searchDepth=1, Name='結束全螢幕')
                            if control.Exists():
                                control.GetInvokePattern().Invoke()
                                auto.Logger.WriteLine(f"{datetime.datetime.today()}|Exit Fullscreen", consoleColor=auto.ConsoleColor.Yellow)
                        # 關閉最大化模式
                        if window_target.IsMaximize():
                            control = window_target.ButtonControl(searchDepth=1, Name = "Restore")
                            control.GetInvokePattern().Invoke()
                            auto.Logger.WriteLine(f"{datetime.datetime.today()}|Exit Maximize", consoleColor=auto.ConsoleColor.Yellow)
                        # 如果視窗在前景 => 取消放大避免擋住其它視窗
                        if window_target.NativeWindowHandle == auto.GetForegroundWindow() and RESIZE_ALREADY == False:
                            RESIZE_X = 1920
                            RESIZE_Y = 1080
                            window_target.GetTransformPattern().Resize(RESIZE_X, RESIZE_Y)
                            auto.Logger.WriteLine(f"{datetime.datetime.today()}|Resize to {RESIZE_X} * {RESIZE_Y}", consoleColor=auto.ConsoleColor.Yellow)
                        RESIZE_ALREADY = True
                    time.sleep(SLEEP_AFTER_FOUND)
                elif NOT_FOUND > 1:
                    FOUND = 0
                    auto.Logger.WriteLine(f"{datetime.datetime.today()}|NOT_FOUND for {NOT_FOUND} times", consoleColor=auto.ConsoleColor.Yellow)
                    window = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client")
                    button = window.ButtonControl(AutomationId = 'PrimaryButton')
                    if button.Exists():
                        button.GetInvokePattern().Invoke()
                        NOT_FOUND = 0
                    else:
                        auto.Logger.WriteLine(f"{datetime.datetime.today()}|NO OK Button Exists", consoleColor=auto.ConsoleColor.Red)
                        NOT_FOUND = 0
                else: # 沒找到目標視窗
                    NOT_FOUND = NOT_FOUND + 1
                    FOUND = 0 
                    RESIZE_ALREADY = False
                    window = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client")
                    if window.Exists():
                        control = window.CustomControl(searchDepth=1)
                        auto.Logger.WriteLine(f"{datetime.datetime.today()}|CustomControl:[ClassName={control.ClassName}] [Name={control.Name}]", consoleColor=auto.ConsoleColor.Yellow)
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
                                list_item = list_control.ListItemControl(searchDepth=1) # TODO 應該加註"vde.vghtpe.gov.tw" or "vdt2.vghtpe.gov.tw"?
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
                            elif control.ClassName == 'LoggingOutView': # 工作階段已逾時的狀態
                                window = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client")
                                window.ButtonControl(searchDepth=1, AutomationId='PrimaryButton').GetInvokePattern().Invoke()
                    else:
                        w = auto.WindowControl(searchDepth=1, Name = "VMware Horizon Client Content Dialog")
                        if w.Exists():
                            w.GetWindowPattern().Close()
                        else:
                            auto.Logger.WriteLine(f"{datetime.datetime.today()}|TARGET WINDOW NOT EXISTS, OPENING...", consoleColor=auto.ConsoleColor.Yellow)
                            os.startfile(PROCESS_PATH)
            else:
                auto.Logger.WriteLine(f"{datetime.datetime.today()}|TARGET PROGRAM NOT EXISTS, OPENING...", consoleColor=auto.ConsoleColor.Yellow)
                process_path =  Path(PROCESS_PATH)
                if process_path.exists():
                    os.startfile(PROCESS_PATH)
                else:
                    auto.Logger.WriteLine(f"{datetime.datetime.today()}|TARGET PROGRAM NOT INSTALLED", consoleColor=auto.ConsoleColor.Yellow)
                    desktop_path = Path.home() / 'Desktop'
                    installer_path = Path(INSTALLER_PATH)
                    dst = shutil.copy2(installer_path, desktop_path)
                    os.startfile(dst)
                    time.sleep(2)
                    # 安裝程式執行
                    installer = auto.WindowControl(searchDepth=1, Name = "Horizon Client 安裝程式")
                    installer.CustomControl(searchDepth=1, ClassName = 'WelcomeView').ButtonControl(searchDepth=1, Name='同意並安裝').GetInvokePattern().Invoke()
                    while True:
                        target = installer.CustomControl(searchDepth=1, ClassName = 'FinishView').ButtonControl(searchDepth=1, Name='完成')
                        if target.Exists():
                            auto.Logger.WriteLine(f"{datetime.datetime.today()}|INSTALLING FINISHED", consoleColor=auto.ConsoleColor.Yellow)
                            break
                    installer.GetWindowPattern().Close()
        except:
            auto.Logger.WriteLine(f"{datetime.datetime.today()}|{traceback.print_exc()}", consoleColor=auto.ConsoleColor.Red)

config = configparser.ConfigParser()
p = Path().glob('autoVMware*.ini')
path_list = list(p)
if len(path_list) == 0: # 處理找不到ini file
    auto.Logger.WriteLine(f"No 'autoVMware*.ini' file", consoleColor=auto.ConsoleColor.Red)
    p_input = input("Please input location of ini file: ")
    path_list.append(Path(p_input))

for p in path_list: # iterate and read the ini file
    try:
        auto.Logger.WriteLine(f"Reading {p}", consoleColor=auto.ConsoleColor.Yellow)
        config.read(p.absolute(), encoding='utf-8-sig') # deal with BOM
        break
    except:
        auto.Logger.WriteLine(f"Reading {p} failed", consoleColor=auto.ConsoleColor.Red)
        continue

ctypes.windll.kernel32.GetUserDefaultUILanguage()

FOUND = 0
NOT_FOUND = 0
RESIZE_ALREADY = False

ENGLISH_VERSION = True if (ctypes.windll.kernel32.GetUserDefaultUILanguage() == 1033) else False # 語言非英文目前視為中文
PROCESS_NAME = config['DEFAULT'].get('PROCESS_NAME')
PROCESS_PATH = config['DEFAULT'].get('PROCESS_PATH', raw=True)
INSTALLER_PATH = config['DEFAULT'].get('INSTALLER_PATH', raw=True)
# INSTALLER_URL = config['DEFAULT'].get('INSTALLER_URL')
WINDOW_TARGET = config['DEFAULT'].get('WINDOW_TARGET')
ACCOUNT = config['DEFAULT'].get('ACCOUNT')
PASSWORD = config['DEFAULT'].get('PASSWORD')
SLEEP_AFTER_FOUND = int(config['DEFAULT'].get('SLEEP_AFTER_FOUND'))
SLEEP_BEFORE_FOUND = int(config['DEFAULT'].get('SLEEP_BEFORE_FOUND'))

auto.Logger.WriteLine(f"PROCESS_NAME:{PROCESS_NAME}\nPROCESS_PATH:{PROCESS_PATH}\nWINDOW_TARGET:{WINDOW_TARGET}", consoleColor=auto.ConsoleColor.Yellow)


if __name__ == '__main__':
    if auto.IsUserAnAdmin():
        main()
    else:
        print('RunScriptAsAdmin', sys.executable, sys.argv)
        auto.RunScriptAsAdmin(sys.argv)