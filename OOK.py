
import win32api
import win32con
import win32gui
import pyperclip
import pyautogui
import win32com.client
import ctypes
import win32clipboard as cb
from pymouse import PyMouse			# 模拟鼠标
from pykeyboard import PyKeyboard	# 模拟键盘
from airtest.core.api import *
from time import sleep
from pynput.keyboard import Controller
import pymysql
import re
import cv2  # https://www.lfd.uci.edu/~gohlke/pythonlibs/#opencv
import numpy as np
from PIL import Image
import operator
from selenium import webdriver
from helium import *

import base64
import uuid

pngTag = 'data:image/png;base64'
jpgTag = 'data:image/jpg;base64'
gifTag = 'data:image/gif;base64'


KEYMAP = {

    "esc": 0x1B,  "window": 0x5B,
    "control": 0x11,    "alt": 0x12,  "kor_eng": 0x15,
    "print_screen": 0x2C,    "scroll_lock": 0x91,   "pause_break": 0x13,


    "f1": 0x70,    "f2": 0x71,    "f3": 0x72,    "f4": 0x73,
    "f5": 0x74,    "f6": 0x75,    "f7": 0x76,    "f8": 0x77,
    "f9": 0x78,    "f10": 0x79,    "f11": 0x7A,    "f12": 0x7B,


    "left_arrow": 0x25,    "right_arrow": 0x27,
    "up_arrow": 0x26,    "down_arrow": 0x28,


    "insert": 0x2D,    "home": 0x24,    "page_up": 0x21,
    "delete": 0x2E,    "end": 0x23,     "page_down": 0x22,


    "backspace": 0x08,  "enter": 0x0D,  "shift": 0x10,
    "tab": 0x09,    "caps_lock": 0x14,  "spacebar": 0x20,


    "0": 0x30,    "1": 0x31,    "2": 0x32,    "3": 0x33,    "4": 0x34,
    "5": 0x35,    "6": 0x36,    "7": 0x37,    "8": 0x38,    "9": 0x39,


    "a": 0x41,    "b": 0x42,    "c": 0x43,    "d": 0x44,    "e": 0x45,
    "f": 0x46,    "g": 0x47,    "h": 0x48,    "i": 0x49,    "j": 0x4A,
    "k": 0x4B,    "l": 0x4C,    "m": 0x4D,    "n": 0x4E,    "o": 0x4F,
    "p": 0x50,    "q": 0x51,    "r": 0x52,    "s": 0x53,    "t": 0x54,
    "u": 0x55,    "v": 0x56,    "w": 0x57,    "x": 0x58,    "y": 0x59,  "z": 0x5A,


    ";": 0xBA,    "=": 0xBB,    ",": 0xBC,    "-": 0xBD,    ".": 0xBE,
    "/": 0xBF,    "`": 0xC0,    "[": 0xDB,    "\\": 0xDC,    "]": 0xDD,
    "'": 0xDE,


    "num_lock": 0x90, "numpad_/": 0x6F, "numpad_*": 0x6A,
    "numpad_-": 0x6D, "numpad_+": 0x6B, "numpad_.": 0x6E,
    "numpad_7": 0x67, "numpad_8": 0x68, "numpad_9": 0x69,
    "numpad_4": 0x64, "numpad_5": 0x65, "numpad_6": 0x66,
    "numpad_1": 0x61, "numpad_2": 0x62, "numpad_3": 0x63,
    "numpad_0": 0x60,
}



UPPER_SPECIAL = {
    "!": 1,    "@": 2,    "#": 3,    "$": 4,    "%": 5,    "^": 6,
    "&": 7,    "*": 8,    "(": 9,    ")": 0,    "_": "-",   "~": '`',    "|": '\\',
    "{": "[",   "}": "]",    ":": ";",    '"': "'", "?": "/", "<": ",", ">": "."
}



def move_mouse(location):
    win32api.SetCursorPos(location)



def get_mouse_position():
    return win32gui.GetCursorPos()



def click(location):
    move_mouse(location)
    l_click()



def right_click(location):
    move_mouse(location)
    r_click()



def double_click(location):
    move_mouse(location)
    l_click()
    l_click()



def key_press_once(key):
    key_on(key)
    key_off(key)



def type_in(string):
    pyperclip.copy(string)
    ctrl_v()



def typing(string):
    for el in string:
        if el.isupper():
            key_on("shift")
            key_press_once(el.lower())
            key_off("shift")
        elif el in UPPER_SPECIAL:
            key_on("shift")
            key_press_once(UPPER_SPECIAL[el])
            key_off("shift")
        else:
            key_press_once(el)



def key_on(key):
    global KEYMAP
    key = str(key)
    if key.isupper:
        key = key.lower()
    try:
        key_code = KEYMAP[key.lower()]
        win32api.keybd_event(key_code, 0, 0x00, 0)
    except KeyError:
        print(key + " is not an available key input.")
        exit(1)



def key_off(key):
    global KEYMAP
    key = str(key)
    if key.isupper:
        key = key.lower()
    try:
        key_code = KEYMAP[key.lower()]
        win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
    except KeyError:
        print(key + " is not an available key input.")
        exit(1)



def l_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)



def r_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)



def mouse_upscroll(number=1000):
    x, y = get_mouse_position()
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, x, y, number, 0)


def mouse_downscroll(number=1000):
    x, y = get_mouse_position()
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, x, y, -1*number, 0)



def drag_drop(frm, to):

    x1, y1 = frm
    x2, y2 = to

    #move_mouse(frm)
    win32api.SetCursorPos(frm)

    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    
    while(x1<x2):
        x1 = x1+2
        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 2, 0, 0, 0)
        time.sleep(0.1)
    #

    #win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE + win32con.MOUSEEVENTF_MOVE, mw, mh, 0, 0)
    

    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)



def get_color(location):
    x, y = location
    return hex(win32gui.GetPixel(win32gui.GetDC(win32gui.GetActiveWindow()), x, y))



def ctrl_c():
    key_on("control")
    key_on("c")
    key_off("control")
    key_off("c")



def ctrl_v():
    key_on("control")
    key_on("v")
    key_off("control")
    key_off("v")



def ctrl_a():
    key_on("control")
    key_on("a")
    key_off("a")
    key_off("control")
    
#notepad++ 新建文档
def ctrl_n():
    key_on("control")
    key_on("n")
    
    key_off("n")
    key_off("control")

#notepad++ 保存文件
def ctrl_s():
    key_on("control")
    key_on("s")
    
    key_off("s")
    key_off("control")
    

def ctrl_f():
    key_on("control")
    key_on("f")
    key_off("control")
    key_off("f")



def alt_f4():
    key_on("alt")
    key_on("f4")
    key_off("alt")
    key_off("f4")



def alt_tab():
    key_on("alt")
    key_on("tab")
    key_off("alt")
    key_off("tab")

def shift_end():
    key_on("shift")
    key_on("end")
    key_off("shift")
    key_off("end")


# 显示桌面
def show_desktop():
    pyautogui.keyDown('winleft')
    pyautogui.press('d')
    pyautogui.keyUp('winleft')
   
# 最大化窗口，以下几行代码都可最大化窗口   
def window_max(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.ShowWindow(hwnd, win32con.SHOW_FULLSCREEN)
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)

# 最小化窗口，以下几行代码都可最大化窗口
def window_min(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
    win32gui.ShowWindow(hwnd, win32con.SW_FORCEMINIMIZE)

def get_clipboard_text():
  cb.OpenClipboard()
  d = cb.GetClipboardData(win32con.CF_TEXT)
  cb.CloseClipboard()
  return d
 
 
def set_clipboard_text(aString):
  cb.OpenClipboard()
  cb.EmptyClipboard()
  cb.SetClipboardData(win32con.CF_TEXT, aString)
  cb.CloseClipboard()


  # 读取剪贴板
def getClipboardText():
    # 打开剪贴板
    cb.OpenClipboard()
    # 读取剪贴板中的数据
    d = cb.GetClipboardData(win32con.CF_UNICODETEXT)
    # 关闭剪贴板
    cb.CloseClipboard()
    # 将读取的数据返回，提供给调用者
    return d
 
  # 设置剪贴板内容
def setClipboardText(aString):
    # 打开剪贴板
    cb.OpenClipboard()
    # 清空剪贴板
    cb.EmptyClipboard()
    # 将数据astring写入剪贴板中
    cb.SetClipboardData(win32con.CF_UNICODETEXT,aString)
    # 关闭剪贴板
    cb.CloseClipboard()

def login(username, password):
    print("")

def getImageFileFromClipboard(tag):
    base64_image = getClipboardText()
    if (jpgTag in base64_image):
        base64_image_without_head = base64_image.replace('data:image/jpg;base64,', '')
        ext = ".jpg"

    if (pngTag in base64_image):
        base64_image_without_head = base64_image.replace('data:image/png;base64,', '')
        ext = ".png"
    
    if (gifTag in base64_image):
        base64_image_without_head = base64_image.replace('data:image/gif;base64,', '')
        ext = ".gif"
        
    result = re.search("data:image/(?P<ext>.*?);base64,(?P<data>.*)", base64_image, re.DOTALL)
    if result:
        ext  = result.groupdict().get("ext")
        data = result.groupdict().get("data")
        print("文件后缀=", ext)
    else:
        raise Exception("Do not parse!")
        
    img = base64.urlsafe_b64decode(data)
    # 3、二进制文件保存
    filename = "{}{}.{}".format(tag, uuid.uuid4(), ext)
    with open(filename, "wb") as f:
        f.write(img)

    return filename


def getBgImg(js):
    touch(Template(r"./img/clear_err.png"))
    time.sleep(0.1)
    pos = exists(Template(r"./img/loc.png"))
    centx,centy = pos
    #click((centx, centy+100))
    
    if(pos):
        click((centx+322, centy+100))
        time.sleep(0.1)
        print('centx=', centx, ', centy=', centy)
    
    #获取背景图
    setClipboardText(js)
    ctrl_v()
    time.sleep(1)
    key_on('enter')
    key_off('enter')
    
    time.sleep(1)
    click((centx+322, centy+100))
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)
    mouse_upscroll(20000)
    time.sleep(0.1)

    pos = exists(Template(r"./img/locBG.png"))
    if(pos):
        x,y = pos
        move_mouse((x, y+30))
        #右键菜单赋值背景图字符串
        r_click()
        #复制
        touch(Template(r"./img/cp_img_str.png"))
        
        return getImageFileFromClipboard('bg')
    else:
        print("获取背景图片失败咯")
        return False

    
def getSlideImg(js):
    touch(Template(r"./img/clear_err.png"))
    
    pos = exists(Template(r"./img/loc.png"))
    x,y = pos
    click((centx, y+100))
    
    #获取背景图
    setClipboardText(js)
    ctrl_v()
    key_on('enter')
    key_off('enter')
    
    #右键菜单copy滑块图字符串
    
    pos = exists(Template(r"./img/locSlider.png"))
    if(pos):
        x,y = pos
        move_mouse((x, y+30))

        r_click()
        #复制
        touch(Template(r"./img/cp_img_str.png"))
        
        return getImageFileFromClipboard('cut')
    else:
        return False

#计算缺口
def getLoc(bgFile, cutFile):
    # 读取背景图片和缺口图片
    bg_img = cv2.imread(bgFile)  # 背景图片
    cut_img = cv2.imread(cutFile)  # 缺口图片
    
    # 识别图片边缘
    bg_edge = cv2.Canny(bg_img, 100, 200)
    cut_edge = cv2.Canny(cut_img, 100, 200)
    
    # 转换图片格式
    bg_pic = cv2.cvtColor(bg_edge, cv2.COLOR_GRAY2RGB)
    cut_pic = cv2.cvtColor(cut_edge, cv2.COLOR_GRAY2RGB)
    
    # 缺口匹配
    res = cv2.matchTemplate(bg_pic, cut_pic, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)  # 寻找最优匹配
    
    # 返回缺口坐标
    return max_loc


def closeConsole():
    pos = exists(Template(r"./img/close_console.png"))
    if(pos):
        print(pos)
    else:
        print("找不到关闭按钮，程序退出")
        exit(0)

    x,y = pos
    click((x,y-12))




if __name__ == '__main__':
    auto_setup(__file__,devices=["Windows:///"])
    me = PyMouse()          # 操作鼠标
    kb = PyKeyboard()       # 操作键盘
    keyboard = Controller()

    #获取背景图的JS代码
    imgBg  = "document.getElementsByClassName('imgBg')[0].style.backgroundImage.replace('url(\"', \"\").replace('\")', \"\")"
    
    #获取滑块图的JS代码
    imgCut = "document.getElementsByClassName('imgBtn')[0].firstChild.src"
    

    
    site_url = 'https://你的网站登录页/admin-techeco#/login'
    username = '你的账号'
    password = '你的密码'

    screen_width = win32api.GetSystemMetrics(0)
    screen_height = win32api.GetSystemMetrics(1)
    print(f"屏幕分辨率为：{screen_width} x {screen_height}")

    centx=(int)(screen_width/2)
    centy=(int)(screen_height/2)
    
    #关闭大写字母锁定键
    if (win32api.GetKeyState(win32con.VK_CAPITAL)):
        key_on('caps_lock')
        key_off('caps_lock')
    
    
    pos = exists(Template(r"./img/lrr.png"))
    if(pos):
        print(pos)
    else:
        touch(Template(r"./img/icon.png"))
        pos = exists(Template(r"./img/lrr.png"))
        if(pos):
            print(pos)
        else:
            print("找不到浏览器，程序退出")
            exit(0)

    hwnd = win32gui.GetForegroundWindow()
    result = win32api.SendMessage(hwnd,win32con.WM_INPUTLANGCHANGEREQUEST,0,0x0409)
    if (result == 0):
        print('设置英文键盘成功！')

    # 输入用户名和密码，点击登录
    x,y = pos
    time.sleep(1)
    click((x+300, y))
    time.sleep(1)
    ctrl_a()
    time.sleep(1)
    key_on('backspace')
    key_off('backspace')
    keyboard.type(site_url)
    time.sleep(1)
    key_on('enter')
    key_off('enter')
    
    time.sleep(2)
    
    #点击登录
    touch(Template(r"./img/login.png"))
    time.sleep(1)
    key_on('f12')
    key_off('f12')
    touch(Template(r"./img/console.png"))
    
    global X2
    global Y2
    # 识别图片边缘   
    bgFile = getBgImg(imgBg)
    siFile = getSlideImg(imgCut)
    if (bgFile and siFile):
        X2, Y2 = getLoc(bgFile, siFile)
        if(X2>160):
            X2 = X2+58
        else:
            X2 = X2+51
        print("文件名：bgFile=", bgFile, ", siFile=", siFile)
    else:
        print("图片文件保存失败")
        exit(1)

    #关闭控制台
    
    
    while(True):
        closeConsole()
        pos = exists(Template(r"./img/drag.png"))
        if (pos):
            X1,Y1 = pos
            X1 = X1-150+20
            print("输出FROME坐标：X=", X1, ", Y=", Y1)
            print("输出TO坐标：X=", X2, ", Y=", Y2)
            
            
            drag_drop((X1,Y1), (X1+X2, Y1))
            print('拖拽结束, from:X1=', X1, ",Y1=", Y1, ", to:X2=", X1+X2, ",Y2=", Y1)
            time.sleep(2)
            pos = exists(Template(r"./img/drag.png"))
            #拖拽失败
            if (pos):
                key_on('f12')
                key_off('f12')
                
                touch(Template(r"./img/console.png"))
                # 识别图片边缘   
                bgFile = getBgImg(imgBg)
                siFile = getSlideImg(imgCut)
                if (bgFile and siFile):
                    X2, Y2 = getLoc(bgFile, siFile)
                    X2 = X2+51
                else:
                    print('文件保存异常')
            else:
                break
        else:
            print('找不到推拽滑块的按钮')
        
    print('程序正常结束..................')
