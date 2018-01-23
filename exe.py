import win32gui
import time
import win32api
import win32con

try:
    classname = "TExamMainForm"
    titlename = "正在考试..."
    #获取父句柄
    hwnd = win32gui.FindWindow(classname, titlename)
    #调整窗口大小
    win32gui.MoveWindow(hwnd, 0, 0, 1000, 550, True)
    #获取所有子句柄
except:
    print("未检测到考试窗口，请重启考试系统或联系作者")
    exit(-2)
while 1:
    try:
        hwndChildList = []
        win32gui.EnumChildWindows(hwnd, lambda hwnd, param: param.append(hwnd),  hwndChildList)
        #print(hwndChildList)
        #列表第十一/十三个为题目，取其标题
        #经观察“首 题”的下一个元素为题目
        for i in range(0,len(hwndChildList)):
            title = win32gui.GetWindowText(hwndChildList[i])
            if title == "首 题":
                temp = win32gui.GetWindowText(hwndChildList[i+1])
                num = temp.split("、",1)[0]
                question = temp.split("、",1)[1]
                print(str(num)+"、"+question)
                #“错”的位置
                #win32api.SetCursorPos([116, 424])
                #“对”的位置
                #win32api.SetCursorPos([116,324])
                #“下题"的位置
                #win32api.SetCursorPos([366,524])
                #"D"的位置
                #win32api.SetCursorPos([116,444])
                #"C"的位置
                #win32api.SetCursorPos([116,394])
                #"B"的位置
                #win32api.SetCursorPos([116,344])
                #"A"的位置
                win32api.SetCursorPos([116,294])
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(100)
    except:
        print("未检测到考试窗口，请重启考试系统或联系作者")
        exit(-1)