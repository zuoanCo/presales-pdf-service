"""测试 PowerPoint COM 状态，带重试，诊断 RETRYLATER 错误。"""
import time

import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = None
for attempt in range(1, 11):
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        print(f"第 {attempt} 次尝试: DispatchEx 成功, version={app.Version}")
        break
    except pythoncom.com_error as e:
        print(f"第 {attempt} 次尝试失败: {e}")
        time.sleep(3)
if app is not None:
    try:
        print("尝试打开演示文稿...")
        pres = app.Presentations.Open(
            r"C:\Users\15944\Desktop\项目\模板HTML\ppt-pdf-service\templates\售前模板.pptx",
            True, False, False,
        )
        print(f"打开成功, 页数={pres.Slides.Count}")
        pres.Close()
    except pythoncom.com_error as e:
        print(f"Open 失败: {e}")
    app.Quit()
pythoncom.CoUninitialize()
