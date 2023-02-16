import threading
from datetime import datetime
import win32con
from pynput import keyboard
from pywinauto.application import Application
from pywinauto import mouse, ElementNotFoundError
import win32api
import psutil
from loguru import logger
from pynput.keyboard import Listener
import time
from pywinauto.controls.uia_controls import ListItemWrapper
from pywinauto.timings import Timings

Timings.fast()

run_status = True
listen_start = True
run_time_start = None
run_time_end = None


def on_press(key):
    global run_status, listen_start
    if isinstance(key, keyboard.KeyCode) and key.char == 'e':
        if listen_start:
            run_status = False
            exit()


def start_listen():
    with Listener(on_press=on_press) as listener:
        listener.join()


th = threading.Thread(target=start_listen, )
th.start()


# from loguru import logger

def next(friend_circle_window_rectangle, wheel_dist=-4):
    logger.info('下滑开始')
    win32api.SetCursorPos(
        (friend_circle_window_rectangle.left + 10, friend_circle_window_rectangle.top + 100)
    )

    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    # time.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    # win32api.keybd_event(40, 0, 0, 0)  # a按下
    # time.sleep(0.1)
    # win32api.keybd_event(40, 0, 0, 0)  # a抬起
    mouse.scroll((friend_circle_window_rectangle.left + 10, friend_circle_window_rectangle.top + 100),
                 wheel_dist=wheel_dist)
    # time.sleep(0.05)
    logger.info('下滑结束')


def time_check():
    global run_time_start, run_time_end, listen_start
    if run_time_start and run_time_end:
        now = datetime.now()
        # 还没到开始时间就延时等待
        if now < run_time_start:
            listen_start = False
            s = int(run_time_start.timestamp() - now.timestamp())
            print(f'未到开始时间，预计等待 {s} 秒')
            time.sleep(s)
            listen_start = True
        # 到时间了就关闭
        if now > run_time_end:
            exit()


def click(rectangle):
    logger.info('准备点击')
    win32api.SetCursorPos((
        rectangle.left + (rectangle.right - rectangle.left) // 2,
        rectangle.top + (rectangle.bottom - rectangle.top) // 2
    ))

    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    # time.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    logger.info('点击完成')


def click_like(app):
    time_check()

    # 刷新朋友圈
    logger.info('获取朋友圈窗口')
    friend_circle_window = app.window(class_name='SnsWnd', found_index=0)
    logger.info('朋友圈窗口获取完成， 开始获取刷新按钮')
    refresh_button = friend_circle_window.child_window(title="刷新", control_type="Button", found_index=0)
    logger.info('刷新按钮获取完成，执行点击')
    refresh_button.click_input()
    click(refresh_button.rectangle())
    logger.info('刷新按钮执行完成')

    liked_num = 0
    while liked_num < 5 and run_status:
        time_check()
        # 获取朋友圈窗口
        friend_circle_window = app.window(class_name='SnsWnd', found_index=0)
        friend_circle_window_rectangle = friend_circle_window.rectangle()
        next(friend_circle_window_rectangle, -3)
        # 获取朋友圈列表
        logger.info('获取朋友圈列表')
        content_items = friend_circle_window.child_window(title="朋友圈", control_type="List", found_index=0)
        logger.info('列表获取完成')
        for content_item in content_items:
            time_check()
            if isinstance(content_item, ListItemWrapper):
                # print(content_item.element_info.name)
                try:
                    # 拿不到这条动态坐标需要下翻
                    logger.info('获取列表中的内容')
                    content = friend_circle_window.child_window(title=content_item.element_info.name, control_type='ListItem', found_index=0)
                    logger.info('内容获取完成，获取评论按钮')
                    content.wait(wait_for='visible', timeout=0.01)
                    comment_button = content.child_window(title="评论", control_type="Button")
                    comment_button.wait(wait_for='visible', timeout=0.01)
                    # comment_button.wait(wait_for='ready', timeout=0.01)
                    # comment_button.wait(wait_for='active', timeout=0.01)
                    logger.info('评论按钮获取完成，获取按钮坐标')
                    comment_button_rectangle = comment_button.rectangle()
                    logger.info('按钮坐标获取完成')
                except Exception:
                    next(friend_circle_window_rectangle)
                    break

                # 顶部工具栏高度 50
                logger.info('执行预检查处理')
                if comment_button_rectangle.top - friend_circle_window_rectangle.top < 30:
                    continue
                # TODO 判断 内容矩形是否再朋友圈范围中 content.rectangle()
                if comment_button_rectangle.bottom > friend_circle_window_rectangle.bottom:
                    next(friend_circle_window_rectangle)
                    break
                logger.info('预检查处理完成，准备点击评论按钮')
                click(comment_button.rectangle())
                # comment_button.click_input()
                logger.info('评论按钮点击完成， 获取评论弹窗')
                # 调试使用

                comment_toast = friend_circle_window.window(class_name='SnsLikeToastWnd', found_index=0)
                # cancel_like_button = comment_toast.child_window(title='取消', control_type="Button")
                logger.info('获取评论弹窗完成')
                print(' - ' * 15)
                try:
                    like_button = comment_toast.child_window(title='赞', control_type="Button")
                    like_button.wait(wait_for='visible', timeout=0.01)
                    # like_button.click_input()
                    print('点赞')
                except ElementNotFoundError:
                    liked_num += 1
                    print('无需点赞')
                print(' - ' * 15)
    print('第一轮点赞完毕，关闭朋友圈')

if __name__ == '__main__':

    run_type = 0
    while True:
        input_txt = input('选择运行时间：\n1-即时，2-时间范围 \n\n\t输入 1 或 2 \n')
        if input_txt in ['1', '2']:
            run_type = input_txt
            break
        print('输入错误')
    today = datetime.now()
    if run_type == '2':
        while True:
            run_time_txt = input('输入开始时间（格式：08:55）：\n')
            try:
                hour, minute = run_time_txt.replace('：', ':').split(':')
                run_time_start = today.replace(hour=int(hour), minute=int(minute), second=0)
                break
            except Exception:
                print('输入错误')

        while True:
            run_time_txt = input('输入结束时间（格式：17:05）：\n')
            try:
                hour, minute = run_time_txt.replace('：', ':').split(':')
                run_time_end = today.replace(hour=int(hour), minute=int(minute), second=0)
                break
            except Exception:
                print('输入错误')

        print(f'程序将于{run_time_start.strftime("%Y-%m-%d %H:%M:%S")}开始，{run_time_end.strftime("%Y-%m-%d %H:%M:%S")}停止')

    wechat_pid_list = []
    process_list = psutil.process_iter()
    for item in process_list:
        if item.name() == 'WeChat.exe':
            wechat_pid_list.append(item.pid)

    index = 0
    while run_status:

        if len(wechat_pid_list) == 0:
            print('需要先登录微信')
            exit()

        time_check()

        if index >= len(wechat_pid_list):
            index = 0
        logger.info('开始自动点赞')
        app = Application(backend='uia').connect(process=wechat_pid_list[index])
        wechat = app.window(class_name='WeChatMainWndForPC')
        # 窗口置顶
        wechat.restore().set_focus()
        # # 打开朋友圈
        logger.info('获取朋友圈按钮')
        friend_circle_button = wechat.child_window(title="朋友圈", control_type="Button", found_index=0)
        logger.info('获取朋友圈按钮完成')
        logger.info('执行点击')
        click(friend_circle_button.rectangle())
        # friend_circle_button.click_input()
        logger.info('点击完成')
        wechat.minimize()
        # 获取朋友圈窗口
        friend_circle_window = app.window(class_name='SnsWnd', found_index=0)
        click_like(app)
        index += 1
