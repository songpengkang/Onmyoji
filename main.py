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
SetWindowPos = windll.user32.SetWindowPos
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0X0002
SWP_NOZORDER = 0x0004

# 图片路径
# img_kaishi = './img/huodong.png'
img_kaishi = './img/huntu.png'   # 魂土
img_jixu = './img/jixu.png'
img_shengli = './img/shengli.png'
yuzhi_kaishi = 0.95  # 游戏开始时，检测到的最大阈值
yuzhi_shengli = 0.93  # 游戏胜利时，检测到的最大阈值
yuzhi_jixu = 0.95    # 游戏继续挑战时，检测到的最大阈值
jiancecishu = 0      # 设置检测次数，大于多少次则停止
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
    byte_array = c_ubyte * total_bytes
    GetBitmapBits(bitmap, total_bytes, byte_array.from_buffer(buffer))
    DeleteObject(bitmap)
    DeleteObject(cdc)
    ReleaseDC(handle, dc)
    # 返回截图数据为numpy.ndarray
    return np.frombuffer(buffer, dtype=np.uint8).reshape(height, width, 4)


def resize_window(handle: HWND, width: int, height: int):
    """设置窗口大小为width × height

    Args:
        handle (HWND): 窗口句柄
        width (int): 宽
        height (int): 高
    """
    SetWindowPos(handle, 0, 0, 0, width, height, SWP_NOMOVE | SWP_NOZORDER)


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
    return zuobiao_x, zuobiao_y, max_val


if __name__ == "__main__":
    if not windll.shell32.IsUserAnAdmin():
        # 不是管理员就提权
        windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, __file__, None, 1)
    handle = windll.user32.FindWindowW(None, "阴阳师-网易游戏")
    resize_window(handle, 1136, 670)
    # 游戏开始界面
    while True:
        image = capture(handle)
        # 转为灰度图
        game = cv2.cvtColor(image, cv2.COLOR_BGRA2GRAY)
        game_kaishi = youxi(game, img_kaishi)      # 检测挑战页面
        game_shengli = youxi(game, img_shengli)    # 检测到胜利页面
        game_jixu = youxi(game, img_jixu)          # 检测到结算页面
        time.sleep(1)
        jiancecishu += 1
        zhongzhi = 0          # 设置终止条件
        print(f'持续检测中，此时游戏开始的最大阈值是{game_kaishi[2]}，游戏继续的最大阈值是{game_jixu[2]}，这是第{jiancecishu}次检测')
        if game_kaishi[2] > yuzhi_kaishi:
            print('检测到游戏开始页面')
            i = 0
            jiancecishu = 0  # 重置检测次数
            while game_kaishi[2] > yuzhi_kaishi:   # 需要一直点击准备，如果只执行一次，在队友不准备的情况下，就会跳过这个步骤。
                left_down(handle, game_kaishi[0], game_kaishi[1])
                time.sleep(random.uniform(0.5, 0.9))
                left_up(handle, game_kaishi[0], game_kaishi[1])
                time.sleep(random.uniform(0.5, 0.9))
                i += 1
                zhongzhi = i
                print(game_kaishi[2])
                print(f'正在第{i}次狂点挑战按钮···坐标是x:{game_kaishi[0]},y:{game_kaishi[1]}')
                image = capture(handle)
                game = cv2.cvtColor(image, cv2.COLOR_BGRA2GRAY)
                game_kaishi = youxi(game, img_kaishi)  # 检测挑战页面
                if game_kaishi[2] < yuzhi_kaishi or i > 10:   # 当队友准备完毕，游戏开始以后，检测结果会小于0.94，则break跳过点击挑战按钮
                    print('游戏开始了···')
                    break
        if game_shengli[2] > yuzhi_shengli:
            print('检测到游戏胜利页面')
            k = 0
            jiancecishu = 0  # 重置检测次数
            while game_shengli[2] > yuzhi_shengli:  # 需要一直点击准备，如果只执行一次，在队友不准备的情况下，就会跳过这个步骤。
                left_down(handle, game_shengli[0], game_shengli[1])
                time.sleep(random.uniform(0.5, 0.9))
                left_up(handle, game_shengli[0], game_shengli[1])
                time.sleep(random.uniform(0.5, 0.9))
                k += 1
                zhongzhi = k
                print(game_shengli[2])
                print(f'正在第{k}次狂点游戏胜利按钮···坐标是x:{game_shengli[0]},y:{game_shengli[1]}')
                image = capture(handle)
                game = cv2.cvtColor(image, cv2.COLOR_BGRA2GRAY)
                game_shengli = youxi(game, img_shengli)  # 检测胜利页面
                if game_shengli[2] < yuzhi_shengli or k > 10:  # 当队友准备完毕，游戏开始以后，检测结果会小于0.94，则break跳过点击挑战按钮
                    print('游戏结束了，正在结算···')
                    break
        if game_jixu[2] > yuzhi_jixu:
            print('检测到继续挑战页面')
            j = 0
            jiancecishu = 0  # 重置检测次数
            while game_jixu[2] > yuzhi_jixu:
                left_down(handle, game_jixu[0], game_jixu[1])
                time.sleep(random.uniform(0.3, 0.5))
                left_up(handle, game_jixu[0], game_jixu[1])
                time.sleep(random.uniform(0.3, 0.5))
                j += 1
                zhongji = j
                print(f'正在第{j}次狂点继续继续继续继续···坐标是x:{game_jixu[0]},y:{game_jixu[1]}')
                image = capture(handle)
                game = cv2.cvtColor(image, cv2.COLOR_BGRA2GRAY)
                game_jixu = youxi(game, img_jixu)  # 检测到结算页面
                if game_jixu[2] < yuzhi_jixu or j > 10:  # 如果小于结算的最大阈值，说明结算完了，正在回到开始页面
                    break
        if jiancecishu > 100 or zhongzhi > 9:
            break


