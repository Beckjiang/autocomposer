import pyautogui
import time
import logging
from typing import List, Optional
import pyperclip
import pygetwindow
from abc import ABC, abstractmethod
import platform
import subprocess
import os

ctrl_key = "ctrl"
if platform.system() == "Darwin":
    ctrl_key = "command"

import cv2
import numpy as np
from PIL import Image
from mss import mss


def capture_screen(monitor_id: int = 1):
    """截取指定显示器的内容"""
    with mss() as sct:
        # 获取所有显示器信息
        monitors = sct.monitors
        print("monitors", len(monitors))
        if monitor_id >= len(monitors):
            raise ValueError(f"显示器 {monitor_id} 不存在")

        # 截取指定显示器
        monitor = monitors[monitor_id]
        screenshot = sct.grab(monitor)

        return Image.frombytes("RGB", screenshot.size, screenshot.rgb)

def check_element_exist(image_path: str, confidence: float = 0.8) -> bool:
    """
    检查元素是否存在（使用 OpenCV 进行图像匹配）

    :param image_path: 目标图像的路径
    :param confidence: 匹配置信度阈值
    :return: 是否存在
    """
    # 1. 截取屏幕
    screenshot = capture_screen(monitor_id=1)  # 返回 PIL 图像
    # 创建tmp目录(如果不存在)
    os.makedirs("./tmp", exist_ok=True)
    # 保存截图,文件名包含时间戳
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    screenshot.save(f"./tmp/debug_screenshot_{timestamp}.png")

    # 2. 将 PIL 图像转换为 OpenCV 格式
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)  # PIL 是 RGB，OpenCV 是 BGR

    # 3. 读取目标图像
    template = cv2.imread(image_path, cv2.IMREAD_COLOR)
    if template is None:
        raise ValueError(f"无法读取目标图像: {image_path}")

    # 4. 使用 OpenCV 进行模板匹配
    result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # 5. 判断匹配结果
    if max_val >= confidence:
        print(f"匹配成功，置信度: {max_val}")
        # 计算目标元素的中心坐标
        h, w = template.shape[:2]  # 获取目标图像的高度和宽度
        center_x = max_loc[0] + w // 2  # 中心点 x 坐标
        center_y = max_loc[1] + h // 2  # 中心点 y 坐标
        return True, (center_x, center_y)
    else:
        print(f"匹配失败，置信度: {max_val}")
        return False, None


class BaseUIAutomator(ABC):
    """自动化框架基类"""

    def __init__(self, config: dict):
        self.config = config
        self.logger = logging.getLogger(self.__class__.__name__)
        self.setup_logging()

    def setup_logging(self):
        """配置日志记录"""
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    @abstractmethod
    def focus_window(self):
        """激活目标窗口（需子类实现）"""
        pass

    def safe_click(self, image_path: str, retry: int = 3):
        """安全点击（带重试机制的图像识别点击）"""
        for _ in range(retry):
            try:
                exist, pos = check_element_exist(image_path)
                if exist:
                    pyautogui.click(pos)
                    return True
            except Exception as e:
                self.logger.warning(f"定位图像时出错: {str(e)}")
            time.sleep(1)
        self.logger.error(f"未找到目标图像：{image_path}")
        return False

    def check_element_exist(self, image_path: str, retry: int = 3, interval: int = 2):
        """检查元素是否存在"""
        for _ in range(retry):
            try:
                exist, pos = check_element_exist(image_path)
                if exist:
                    return True
            except Exception as e:
                self.logger.warning(f"元素不存在: {str(e)}")
            time.sleep(interval)
        return False

    def input_text(self, text: str, lang: str = 'en'):
        """多语言安全输入"""
        if lang == 'zh' or platform.system() == "Darwin":  # 对于中文输入或macOS系统
            pyperclip.copy(text)
            time.sleep(0.5)

            # 使用更可靠的粘贴方法
            pyautogui.keyDown(ctrl_key)
            time.sleep(0.5)
            pyautogui.press('v')
            time.sleep(0.5)
            pyautogui.keyUp(ctrl_key)
            time.sleep(0.5)
        else:
            pyautogui.typewrite(text)
        time.sleep(0.5)  # 增加输入后的等待时间

    def _hotkey(self, keys_list: List[str]):
        """改进的组合键实现，不修改原始列表"""
        keys = keys_list.copy()  # 复制列表避免修改原始数据
        first_key = keys.pop(0)

        # 对于macOS，增加按键可靠性
        if platform.system() == "Darwin":
            print(f"执行组合键: {first_key} + {keys}")
            pyautogui.keyDown(first_key)
            time.sleep(0.5)  # 增加按键之间的延时
            for key in keys:
                pyautogui.keyDown(key)
                time.sleep(0.3)
                pyautogui.keyUp(key)
                time.sleep(0.3)
            time.sleep(0.5)
            pyautogui.keyUp(first_key)
            time.sleep(0.5)
        else:
            # 原有实现
            with pyautogui.hold(first_key):
                print("hold", first_key)
                time.sleep(1)
                if len(keys) > 0:
                    for key in keys:
                        pyautogui.press(key)
                        print("press", key)
                        time.sleep(0.5)  # 按键之间添加延时
            print("release", first_key)


class ComposerAutomator(BaseUIAutomator):
    """Composer自动化专用类"""

    def __init__(self, config: dict):
        super().__init__(config)
        self.validate_config()

    def validate_config(self):
        """验证必要配置项"""
        required_keys = ['hotkeys', 'wait_timeouts', 'ui_elements']
        for key in required_keys:
            if key not in self.config:
                raise ValueError(f"Missing required config key: {key}")

    def focus_window(self):
        """激活Composer窗口"""
        try:
            if platform.system() == "Darwin":
                print("在macOS上使用AppleScript激活窗口")
                # 打出所有窗口的title
                print(pygetwindow.getAllTitles())

                # 首先尝试用AppleScript激活应用程序本身
                app_script = '''
                tell application "Cursor"
                    activate
                end tell
                '''
                subprocess.run(["osascript", "-e", app_script], capture_output=True)
                time.sleep(1)

                # 然后再查找特定窗口
                window_title = self.config['window_title']
                window_script = f'''
                tell application "System Events"
                    tell process "Cursor"
                        set frontmost to true
                        set allWindows to every window
                        repeat with aWindow in allWindows
                            if title of aWindow contains "{window_title}" then
                                set visible of aWindow to true
                                perform action "AXRaise" of aWindow
                                return true
                            end if
                        end repeat
                    end tell
                end tell
                '''
                subprocess.run(["osascript", "-e", window_script], capture_output=True)
            else:
                # 在Windows上使用pygetwindow
                win = pygetwindow.getWindowsWithTitle(self.config['window_title'])[0]
                win.activate()

            time.sleep(1.5)  # 增加等待时间
            return True
        except Exception as e:
            self.logger.error(f"激活目标窗口失败: {str(e)}")
            return False

    def open_project(self):
        """打开项目"""
        # 执行命令
        if platform.system() == "Windows":
            # Windows系统下，尝试使用完整路径，或使用shell=True允许系统搜索PATH
            try:
                # 优先尝试直接使用shell=True方式执行
                subprocess.run(f"cursor {self.config['project_path']}", shell=True)
            except Exception as e:
                self.logger.warning(f"使用shell方式打开项目失败: {str(e)}")
                # 尝试查找Cursor可能的安装位置
                cursor_paths = [
                    r"C:\Program Files\Cursor\Cursor.exe",
                    r"C:\Users\%USERNAME%\AppData\Local\Programs\Cursor\Cursor.exe",
                    # 添加其他可能的安装路径...
                ]

                for path in cursor_paths:
                    expanded_path = os.path.expandvars(path)
                    if os.path.exists(expanded_path):
                        subprocess.run([expanded_path, self.config['project_path']])
                        break
                else:
                    self.logger.error("无法找到Cursor可执行文件，请手动设置正确的路径")
        else:
            # 其他系统保持原样
            subprocess.run(["cursor", self.config['project_path']])

        time.sleep(2)

    def open_composer(self):
        """打开Composer"""
        self.logger.info("正在打开Composer...")
        # 先确保窗口焦点
        self.focus_window()

        # 执行热键并等待
        self._hotkey(self.config['hotkeys']['open_composer'])
        time.sleep(self.config['wait_timeouts']['after_open'])

        # 验证是否成功打开
        retry = 3
        while retry > 0:
            if self.check_element_exist(self.config['ui_elements']['agent_mode'], retry=1, interval=1):
                self.logger.info("Composer已成功打开")
                return True
            else:
                retry -= 1
                self.logger.info(f"未检测到Composer界面，重试中 (剩余{retry}次)...")
                self.focus_window()  # 重新获取焦点
                self._hotkey(self.config['hotkeys']['open_composer'])
                time.sleep(self.config['wait_timeouts']['after_open'])

        self.logger.warning("打开Composer失败，可能需要手动干预")
        return False

    def new_session(self):
        """新建会话"""
        self.logger.info("正在新建会话...")

        # 先确保窗口焦点
        self.focus_window()
        time.sleep(1)

        # 执行新建会话热键
        self._hotkey(self.config['hotkeys']['new_session'])
        time.sleep(1)  # 给新会话更多时间加载

        # 再次确保窗口焦点
        self.focus_window()
        time.sleep(1)

        # 清空可能的现有内容，一次性删除所有内容
        pyautogui.hotkey('command', 'a') if platform.system() == "Darwin" else pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.press('delete')
        time.sleep(0.5)
        pyautogui.press('delete')
        time.sleep(0.5)

        self.logger.info("新会话创建完成，输入区域正常")
        return True

    def link_files(self, file_paths: List[str]):
        """关联多个文件"""
        self.logger.info(f"正在关联 {len(file_paths)} 个文件...")

        # 强制获取窗口焦点
        self.focus_window()
        time.sleep(1)


        for path in file_paths:
            self.logger.info(f"正在关联文件: {path}")

            # 使用单独的按键事件替代typewrite
            pyautogui.press(' ')
            time.sleep(0.3)

            pyautogui.press('@')
            time.sleep(0.5)

            # 直接使用命令行参数粘贴
            if platform.system() == "Darwin":
                # macOS特定方法：使用pbpaste命令确认复制内容
                pyperclip.copy(path)
                time.sleep(0.3)
                result = subprocess.run(["pbpaste"], capture_output=True, text=True)
                self.logger.info(f"剪贴板内容: {result.stdout}")

                # 确保使用可靠的粘贴方式
                for attempt in range(1):  # 尝试粘贴多次
                    pyautogui.keyDown('command')
                    time.sleep(0.3)
                    pyautogui.keyDown('v')
                    time.sleep(0.3)
                    pyautogui.keyUp('v')
                    time.sleep(0.3)
                    pyautogui.keyUp('command')
                    time.sleep(0.5)
            else:
                pyperclip.copy(path)
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'v')

            time.sleep(1)

            # 多次尝试按回车确认
            # for _ in range(2):
            pyautogui.press('enter')
            time.sleep(0.5)

            time.sleep(self.config['wait_timeouts']['per_file'])

            # 确认文件引用是否成功
            # 这里可以添加图像识别检测，暂时先依赖等待

    def input_prompt(self, prompt: str):
        """输入业务提示"""
        self.logger.info("输入业务提示...")

        # 输入提示内容
        self.input_text(prompt, lang='zh')
        time.sleep(0.5)

        # 点击发送按钮或使用快捷键发送
        self.send_message()

    def send_message(self):
        """发送消息"""
        self.logger.info("点击发送按钮...")

        # 先尝试通过图像识别找到发送按钮
        if self.safe_click(self.config['ui_elements']['send_button']):
            self.logger.info("成功点击发送按钮")
            return True

        # 如果图像识别失败，则尝试通过快捷键发送
        self.logger.info("尝试使用快捷键发送...")
        if platform.system() == "Darwin":
            pyautogui.hotkey('command', 'enter')  # macOS上的快捷键
        else:
            pyautogui.hotkey('ctrl', 'enter')  # Windows上的快捷键

        time.sleep(0.5)
        return True

    def execute_workflow(self, task: str, files: List[str]):
        """完整工作流执行"""
        try:
            print(pyautogui.size())
            self.logger.info("开始自动化工作流...")

            # 打开项目并等待足够时间加载
            self.open_project()
            time.sleep(3)

            # 确保窗口焦点
            success = self.focus_window()
            if not success:
                self.logger.error("无法获取窗口焦点，工作流中断")
                return False

            # 打开作曲家界面
            self.logger.info("步骤1: 打开Composer界面")
            composer_opened = self.open_composer()
            if not composer_opened:
                self.logger.error("无法打开Composer，工作流中断")
                return False
            time.sleep(1)

            # 创建新会话
            self.logger.info("步骤2: 创建新会话")
            self.new_session()
            time.sleep(1)

            # 关联文件
            self.logger.info("步骤3: 关联文件")
            self.link_files(files)

            # 输入提示
            if task:
                self.logger.info("步骤4: 输入任务提示")
                self.focus_window()
                time.sleep(1)
                self.input_prompt(task)

                # 等待执行完成
                self.logger.info("步骤5: 等待执行完成")
                self.wait_execution()

                # 接受更改
                self.logger.info("步骤6: 接受更改")
                return self.accept_changes()

            return True
        except Exception as e:
            self.logger.error(f"流程执行失败: {str(e)}")
            import traceback

            self.logger.error(traceback.format_exc())
            return False

    def wait_execution(self, timeout: Optional[int] = None):
        """智能等待执行完成"""
        timeout = timeout or self.config['wait_timeouts']['execution']
        start = time.time()

        while time.time() - start < timeout:
            try:
                if self.check_element_exist(self.config['ui_elements']['generating']):
                    time.sleep(5)
                else:
                    self.logger.info("执行已完成")
                    return True
            except Exception as e:
                self.logger.error(f"等待执行完成失败: {str(e)}")
                return False
        raise TimeoutError("执行超时")

    def accept_changes(self):
        """接收所有修改"""
        pyautogui.hotkey(ctrl_key, 'enter')
        time.sleep(0.5)
        return True

# 配置文件示例
CONFIG = {
    "window_title": "fastapi-demo", # 窗口标题
    "project_path": "D:/fastapi-demo", # 项目路径
    "hotkeys": {"open_composer": [ctrl_key, "i"], "new_session": [ctrl_key, "n"]}, # 快捷键
    "wait_timeouts": {"after_open": 1, "per_file": 1, "execution": 300}, # 等待时间
    "ui_elements": {
        "agent_mode": "./images/windows/agent_mode.png",
        "send_button": "./images/windows/send_button.png",
        "generating": "./images/windows/generating.png",
    },
}

# 使用示例
if __name__ == "__main__":
    automator = ComposerAutomator(CONFIG)
    # 完整流程执行
    success = automator.execute_workflow(
        task="初始化一个fastapi项目，构建对于实体Task的curd接口",
        files=[],
    )

    if success:
        print("流程执行成功！")
    else:
        print("流程执行失败！")
