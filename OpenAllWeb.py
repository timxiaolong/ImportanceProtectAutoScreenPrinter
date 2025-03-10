import os
import time
import webbrowser

import pythoncom
import win32com.client
from apscheduler.schedulers.blocking import BlockingScheduler


def job():
    global Prosstimes
    print("任务启动")
    # 获取url表格
    excel_file = os.path.join(current_dir, 'connectUrl.xlsx')
    sheet_name = 'Sheet1'
    # 选取表格链接的区间
    url_range = 'A1:A86'
    open_excel_url_and_screenshot(excel_file, sheet_name, url_range)
    print("任务结束")


def open_excel_url_and_screenshot(excel_file, sheet_name, url_range):
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
            time.sleep(1)

    # 关闭Excel
    workbook.Close(False)  # False表示不保存更改
    excel.Quit()
    # 检测启动是哪个浏览器
    pythoncom.CoUninitialize()

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))
Prosstimes = 1
job()
# scheduler = BlockingScheduler()
# scheduler.add_job(job)
# scheduler.start()


