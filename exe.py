import win32gui
import time
import xlrd
import win32api
import win32con

test = input("请调整cmd窗口到屏幕右侧，按回车键开始自动答题")

data = xlrd.open_workbook('exercise.xls')
table = data.sheets()[0]
nrows = table.nrows
lenOfXls = len(data.sheets())
questionlist = []
for x in range(0, lenOfXls):
    xls = data.sheets()[x]
    for i in range(1, nrows):
        temp = xls.cell(i, 1).value
        questionlist.append(temp)

def searchanswer():
    global answerA, answerB, answerC, answerD
    answerA = None
    answerB = None
    answerC = None
    answerD = None
    num = [i for i, x in enumerate(questionlist) if x == question]
    answer = xls.cell(num[0] + 1, 7).value
    list1 = []
    for item in range(0, len(answer), 1):
        temp = str(answer[item:item + 1])
        list1.append(temp)
    if "A" in list1:
        answerA = xls.cell(num[0] + 1, 3).value
        print(answerA)
    if "B" in list1:
        answerB = xls.cell(num[0] + 1, 4).value
        print(answerB)
    if "C" in list1:
        answerC = xls.cell(num[0] + 1, 5).value
        print(answerC)
    if "D" in list1:
        answerD = xls.cell(num[0] + 1, 6).value
        print(answerD)
    return answerA,answerB,answerC,answerD
def click():
    allanswer = list(searchanswer())
    if A in allanswer:
        win32api.SetCursorPos([113, 298])
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.001)
    if B in allanswer:
        win32api.SetCursorPos([116, 344])
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.01)
    if C in allanswer:
        if C == "对":
            win32api.SetCursorPos([116, 324])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        else:
            win32api.SetCursorPos([116, 394])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0,0)
        time.sleep(0.01)
    if D in allanswer:
        if D == "错":
            win32api.SetCursorPos([116, 424])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        else:
            win32api.SetCursorPos([116, 444])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0,0)
        time.sleep(0.01)
    win32api.SetCursorPos([366, 524])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0,0)


checkquestion = ["test", "test1"] #填充两个无用元素避免越界
#经观察“首 题”的下一个元素为题目
while 1:
    try:
        try:
            classname = "TExamMainForm"
            titlename = "正在考试..."
            # 获取父句柄
            hwnd = win32gui.FindWindow(classname, titlename)
            # 调整窗口大小
            win32gui.MoveWindow(hwnd, 0, 0, 1000, 550, True)
            # 获取所有子句柄
        except:

            print("未检测到考试窗口，请重启考试系统或联系作者")
            exit(-2)
        hwndChildList = []
        win32gui.EnumChildWindows(hwnd, lambda hwnd, param: param.append(hwnd), hwndChildList)

        for i in range(0,len(hwndChildList)):
            title = win32gui.GetWindowText(hwndChildList[i])
            if title == "首 题":
                temp = win32gui.GetWindowText(hwndChildList[i+1])
                num = temp.split("、",1)[0]
                question = temp.split("、",1)[1]
                checkquestion.append(question)
                D = win32gui.GetWindowText(hwndChildList[2])
                C = win32gui.GetWindowText(hwndChildList[3])
                B = win32gui.GetWindowText(hwndChildList[4])
                A = win32gui.GetWindowText(hwndChildList[5])
                print(str(num)+"、"+question) #打印题目题号
                if checkquestion[-1] != checkquestion[-2]:
                    if question in questionlist:
                        click()
                    if question not in questionlist:
                        question = question.replace(",","，")
                        if question in questionlist:
                            click()

        time.sleep(0.001)
    except:
        pass
