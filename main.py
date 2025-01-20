import datetime
import time

import psutil
import pyautogui
import pythoncom
import webbrowser
import win32com.client
import os
from apscheduler.schedulers.blocking import BlockingScheduler

def job():
    global Prosstimes
    print("在"+time.strftime("%Y-%m-%d,%H:%M:%S")+"执行第"+str(Prosstimes)+"任务")
    # 获取url表格
    excel_file = os.path.join(current_dir, 'url.xlsx')
    sheet_name = 'Sheet1'
    url_range = 'A1:A76'
    screenshot_region = (0, 35, 1920, 1045)  # 例如，截取从坐标(100, 200)开始，宽800像素，高600像素的区域
    open_excel_url_and_screenshot(excel_file, sheet_name, url_range, screenshot_region,creat_time_folder())
    print("第"+str(Prosstimes)+"次任务在"+time.strftime("%Y-%m-%d,%H:%M:%S")+"结束，等待运行下一次")
    Prosstimes = Prosstimes + 1


def creat_time_folder():
    next_hour = int(time.strftime("%H"))+1
    forder_name = time.strftime("%Y-%m-%d,%H")+'-'+str(next_hour)+'ScreenShot'
    os.chdir(current_dir+'/ScreenShot')
    os.makedirs(forder_name)
    print("创建新的文件夹，命名为："+forder_name)
    return 'ScreenShot'+'/'+forder_name

def open_excel_url_and_screenshot(excel_file, sheet_name, url_range, screenshot_region, folder_name):
    """
    打开Excel文件，获取指定单元格的URL，并截取指定区域的屏幕截图。

    Args:
      excel_file: Excel文件路径
      sheet_name: 要打开的工作表名称
      url_cell: 包含URL的单元格地址（例如，"A1"）
      screenshot_region: 要截取的屏幕区域，以元组形式表示（左上角x坐标，左上角y坐标，宽度，高度）
    """

    # 打开Excel文件
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # 可选：设置为True显示Excel窗口
    workbook = excel.Workbooks.Open(excel_file)
    worksheet = workbook.Worksheets(sheet_name)
    # 获取URL范围
    url_range = worksheet.Range(url_range)
    flag = 1

    # 遍历每个URL并打开
    for cell in url_range:
        flag_str = str(flag)
        url = cell.Value
        if url:  # 判断单元格是否有值
            webbrowser.open(url)
            # 等待页面加载（可根据实际情况调整等待时间）
            time.sleep(10)
            # 截取屏幕
            screenshot = pyautogui.screenshot(region=screenshot_region)
            file_name = 'screenshot'+flag_str+'.png'
            os.chdir(current_dir+'/'+'/'+folder_name)
            screenshot.save(file_name)  # 保存截图
            flag = flag + 1

    # 关闭Excel
    workbook.Close(False)  # False表示不保存更改
    excel.Quit()
    # 检测启动是哪个浏览器
    pl = psutil.pids()
    try:
        for pid in pl:
            if(psutil.Process(pid).name() == "msedge.exe"):
                os.system('taskkill /F /IM msedge.exe')
            elif(psutil.Process(pid).name() == "opera.exe"):
                os.system('taskkill /F /IM opera.exe')
    except:
        print("浏览器程序已结束")
    pythoncom.CoUninitialize()

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))
Prosstimes = 1
job()
scheduler = BlockingScheduler()
scheduler.add_job(job, 'interval', seconds=30)
scheduler.start()


