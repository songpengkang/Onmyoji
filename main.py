from ctypes import windll, byref, c_ubyte
from ctypes.wintypes import RECT, HWND
import numpy as np
import cv2
import time
import sys
import random


GetDC = windll.user32.GetDC
CreateCompatibleDC = windll.gdi32.CreateCompatibleDC
GetClientRect = windll.user32.GetClientRect
CreateCompatibleBitmap = windll.gdi32.CreateCompatibleBitmap
SelectObject = windll.gdi32.SelectObject
BitBlt = windll.gdi32.BitBlt
SRCCOPY = 0x00CC0020
GetBitmapBits = windll.gdi32.GetBitmapBits
DeleteObject = windll.gdi32.DeleteObject
ReleaseDC = windll.user32.ReleaseDC
PostMessageW = windll.user32.PostMessageW
ClientToScreen = windll.user32.ClientToScreen
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x202

# 图片路径
img_kaishi = './duibi_img/huodong.png'
img_jixu = './duibi_img/jixu.png'
# 排除缩放干扰
windll.user32.SetProcessDPIAware()


def capture(handle: HWND):
    """窗口客户区截图

    Args:
        handle (HWND): 要截图的窗口句柄

    Returns:
        numpy.ndarray: 截图数据
    """
    # 获取窗口客户区的大小
    r = RECT()
    GetClientRect(handle, byref(r))
    width, height = r.right, r.bottom
    # 开始截图
    dc = GetDC(handle)
    cdc = CreateCompatibleDC(dc)
    bitmap = CreateCompatibleBitmap(dc, width, height)
    SelectObject(cdc, bitmap)
    BitBlt(cdc, 0, 0, width, height, dc, 0, 0, SRCCOPY)
    # 截图是BGRA排列，因此总元素个数需要乘以4
    total_bytes = width*height*4
    buffer = bytearray(total_bytes)
    byte_array = c_ubyte*total_bytes
    GetBitmapBits(bitmap, total_bytes, byte_array.from_buffer(buffer))
    DeleteObject(bitmap)
    DeleteObject(cdc)
    ReleaseDC(handle, dc)
    # 返回截图数据为numpy.ndarray
    return np.frombuffer(buffer, dtype=np.uint8).reshape(height, width, 4)


def move_to(handle: HWND, x: int, y: int):
    """移动鼠标到坐标（x, y)

    Args:
        handle (HWND): 窗口句柄
        x (int): 横坐标
        y (int): 纵坐标
    """
    # https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-mousemove
    wparam = 0
    lparam = y << 16 | x
    PostMessageW(handle, WM_MOUSEMOVE, wparam, lparam)


def left_down(handle: HWND, x: int, y: int):
    """在坐标(x, y)按下鼠标左键

    Args:
        handle (HWND): 窗口句柄
        x (int): 横坐标
        y (int): 纵坐标
    """
    # https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-lbuttondown
    wparam = 0
    lparam = y << 16 | x
    PostMessageW(handle, WM_LBUTTONDOWN, wparam, lparam)


def left_up(handle: HWND, x: int, y: int):
    """在坐标(x, y)放开鼠标左键

    Args:
        handle (HWND): 窗口句柄
        x (int): 横坐标
        y (int): 纵坐标
    """
    # https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-lbuttonup
    wparam = 0
    lparam = y << 16 | x
    PostMessageW(handle, WM_LBUTTONUP, wparam, lparam)


def youxi(img_bg,img_path):
    templete_img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
    result = cv2.matchTemplate(img_bg, templete_img, cv2.TM_CCORR_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    top_left = max_loc
    h, w = templete_img.shape[:2]
    bottom_right = top_left[0] + w, top_left[1] + h
    zuobiao_x = random.randint(top_left[0], bottom_right[0])
    zuobiao_y = random.randint(top_left[1], bottom_right[1])
    return zuobiao_x, zuobiao_y, max_val, top_left, bottom_right


if __name__ == "__main__":
    if not windll.shell32.IsUserAnAdmin():
        # 不是管理员就提权
        windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, __file__, None, 1)
    handle = windll.user32.FindWindowW(None, "阴阳师-网易游戏")
    # 游戏开始界面
    while True:
        image = capture(handle)
        # 转为灰度图
        game = cv2.cvtColor(image, cv2.COLOR_BGRA2GRAY)
        game_kaishi = youxi(game, img_kaishi)
        print('游戏等待开始中')
        print(game_kaishi[2])
        if game_kaishi[2] > 0.94:
            time.sleep(random.uniform(1.1,2.2))
            left_down(handle, game_kaishi[0], game_kaishi[1])
            time.sleep(random.uniform(0.5, 1))
            left_up(handle, game_kaishi[0], game_kaishi[1])
            print('游戏已经开始了')
            time.sleep(5)
            while True:
                image = capture(handle)
                game1 = cv2.cvtColor(image, cv2.COLOR_BGRA2GRAY)
                game_jixu = youxi(game1, img_jixu)
                # cv2.rectangle(game1, game_jixu[3], game_jixu[4], (0, 0, 255), 2)
                # cv2.imshow('Match Template', game1)
                # cv2.waitKey()
                print('游戏进行中')
                print(game_jixu[2])
                if game_jixu[2] > 0.99:
                    time.sleep(random.uniform(1.1, 2.2))
                    left_down(handle, game_jixu[0], game_jixu[1])
                    time.sleep(random.uniform(0.5, 1.0))
                    left_up(handle, game_jixu[0], game_jixu[1])
                    print('已经完成了点击游戏继续操作')
                    break





