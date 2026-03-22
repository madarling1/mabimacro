import sys
import os
import time
import threading
import random
import cv2
import numpy as np
import ctypes
import win32gui
import win32ui
import win32con
import win32api
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTextEdit, QLabel, QRadioButton, QGroupBox, 
                             QDialog, QTextBrowser, QLineEdit, QGraphicsDropShadowEffect,
                             QFrame, QButtonGroup)
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, Qt, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import (QFont, QIcon, QColor, QFontDatabase)

# [설정] 콘솔 숨김
kernel32 = ctypes.WinDLL('kernel32')
user32 = ctypes.WinDLL('user32')
hWnd = kernel32.GetConsoleWindow()
if hWnd:
    user32.ShowWindow(hWnd, win32con.SW_HIDE)

# [설정] DPI
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    ctypes.windll.user32.SetProcessDPIAware()


# ── 컬러 팔레트 (파스텔 다크) ──
COLORS = {
    'bg':           '#1e1525',
    'card':         '#2a2035',
    'card_border':  '#4a3860',
    'surface':      '#302540',
    'text':         '#f5eafa',
    'text_dim':     '#a890c0',
    'accent':       '#f48cb6',
    'accent_hover': '#f8a0c8',
    'success':      '#7ed8a6',
    'success_hover':'#90e8b8',
    'danger':       '#f28080',
    'danger_hover': '#f89898',
    'log_bg':       '#160e1e',
    'log_text':     '#daa0f0',
    'input_bg':     '#221a2c',
    'input_border': '#4a3860',
}
class NoticeDialog(QDialog):
    """매크로 시작 전, 사용자가 꼭 알아야 할 주의사항과 가이드를 보여주는 안내창입니다."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI(parent)

    def initUI(self, parent):
        self.setWindowTitle('모비노기 통합 매크로 가이드')
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setFixedSize(400, 560)

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        icon_path = os.path.join(base_path, 'my_icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        if parent:
            self.setGeometry(parent.geometry())

        c = COLORS
        
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {c['bg']};
            }}
        """)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        self.notice_text = QTextBrowser()
        self.notice_text.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {c['card']};
                color: {c['text']};
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                border: 1px solid {c['card_border']};
                border-radius: 10px;
                padding: 12px;
            }}
        """)
        
        content = """
        <h3 style='color: #6c5ce7;'>[ 필독: 매크로 가이드 ]</h3>
        <p>이 매크로는 마력기폭제 재료템 파밍, 자동 낚시를 위해 제작되었습니다. 마력 깃든 돌, 반짝이는 이끼를 제외한 파밍템을 자동으로 판매하고 인벤토리를 확보합니다.</p>
        <p><b>[마력 깃든 돌]</b> : 매크로를 실행하면 자동으로 무한채집이 가능.</p>
        <p><b>[반짝이는 이끼]</b> : 황금 이끼를 사용하여 수동으로 무한채집.</p>
        <p><b>[낚시]</b> : 집중해서 낚시를 선택 후 사용.</p>
        <hr>
        <p>백동 광석 몇개를 캐서, <b>인벤토리에 백동 광석과 돌멩이, 마력 깃든 돌이 있는 상태</b>로 자동채집하며 매크로 실행</p>
        <p>이끼도 튼튼버섯 몇개를 캐서, <b>인벤토리에 튼튼버섯과 튼튼버섯 포자가 있는 상태</b>로 자동채집을 시작한 후 매크로 실행</p> 
        <hr>
        <p><b>1. 작업별 위치</b></p>
        <ul>
            <li>마력 깃든 돌 : 얼음협곡 광맥</li>
            <li>반짝이는 이끼 : 시드스넷타 버섯</li>
            <li>낚시 : 원하는 낚시터 진입 후 앞으로 가도 안 움직이는 자리</li>
        </ul>
        <p><b>2. 매크로 실행 전 필수 마비노기 세팅</b></p>
        <ul>
            <li>반드시 <b>인벤토리를 연 후 하단의 아이템 탭</b>을 클릭하고 닫으세요.</li>
            <li>창 크기: <b>1280x960</b> 설정 후, <b>세로만 최소</b>로 줄여주세요.</li>
        </ul>
        <p><b>3. 주의사항</b></p>
        <ul>
            <li>인벤토리에 마력 깃든 돌이 최소 1개 이상 있어야 합니다.</li>
            <li>이미지 스캔 방식이라 창이 가려지면 안 됩니다.</li>
            <li> 간혹 펫 때문에 낚시 오류가 뜨는 경우가 있으니 <b>펫은 꼭 넣고 진행</b>해주세요.</li>
        </ul>
        """
        self.notice_text.setHtml(content)
        
        self.btn_close = QPushButton("확인했습니다")
        self.btn_close.setFixedHeight(44)
        self.btn_close.setCursor(Qt.PointingHandCursor)
        self.btn_close.setStyleSheet(f"""
            QPushButton {{
                background-color: {c['accent']};
                color: white;
                font-weight: bold;
                font-size: 11pt;
                border: none;
                border-radius: 10px;
            }}
            QPushButton:hover {{
                background-color: {c['accent_hover']};
            }}
        """)
        self.btn_close.clicked.connect(self.accept)
        
        layout.addWidget(self.notice_text)
        layout.addWidget(self.btn_close)
        self.setLayout(layout)


class MacroWorker(QObject):
    """실제로 게임 화면을 분석하고 조작하는 '일꾼' 클래스입니다."""
    log_signal = pyqtSignal(str)
    countdown_signal = pyqtSignal(int)
    finished_signal = pyqtSignal()
    stats_signal = pyqtSignal(str)

    def __init__(self, mode):
        super().__init__()
        self.is_running = False
        self.mode = mode
        self.GAME_TITLE = "마비노기 모바일"
        
        self.fishing_conf = {
            'ready': ['ready.png'],
            'yes': ['yes1.png', 'yes2.png'],
            'no': ['no1.png', 'no2.png'],
            'threshold': 0.8,
            'timeout': 10.0
        }

    def check_running(self):
        if not self.is_running:
            raise InterruptedError("중지됨")

    def interruptible_sleep(self, seconds):
        steps = int(seconds / 0.1)
        for _ in range(steps):
            self.check_running()
            time.sleep(0.1)
        time.sleep(seconds % 0.1)

    def get_window_screenshot(self, hwnd):
        try:
            rect = win32gui.GetWindowRect(hwnd)
            w, h = rect[2] - rect[0], rect[3] - rect[1]
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
            saveDC.SelectObject(saveBitMap)
            ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)
            signedIntsArray = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(signedIntsArray, dtype='uint8')
            img.shape = (h, w, 4)
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        except:
            return None

    def background_click_pro(self, hwnd, x, y):
        self.check_running()
        old_pos = win32api.GetCursorPos()
        rect = win32gui.GetWindowRect(hwnd)
        screen_x = rect[0] + int(x)
        screen_y = rect[1] + int(y)
        try:
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_CLICKACTIVE, 0)
            win32api.SetCursorPos((screen_x, screen_y))
            self.interruptible_sleep(0.05)
            ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            self.interruptible_sleep(0.1)
            ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        finally:
            self.interruptible_sleep(0.05)
            win32api.SetCursorPos(old_pos)

    def send_key_perfect(self, hwnd, vk_code, scan_code, duration=0.1):
        self.check_running()
        win32gui.PostMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        down_lparam = 1 | (scan_code << 16)
        up_lparam = 1 | (scan_code << 16) | (1 << 30) | (1 << 31)
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, down_lparam)
        self.interruptible_sleep(duration)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, up_lparam)

    def find_image_pos(self, screen, image_name, threshold=0.7):
        if screen is None: return None
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(base_path, image_name)
        if not os.path.exists(full_path): return None
        try:
            img_array = np.fromfile(full_path, np.uint8)
            template = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            if template is None: return None
            res = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            if max_val >= threshold:
                th, tw = template.shape[:2]
                return (max_loc[0] + tw // 2, max_loc[1] + th // 2)
        except:
            pass
        return None

    def find_and_click(self, hwnd, image_name, threshold=0.7, wait_time=5, label=None):
        display = label or image_name
        start_time = time.time()
        while time.time() - start_time < wait_time:
            self.check_running()
            screen = self.get_window_screenshot(hwnd)
            pos = self.find_image_pos(screen, image_name, threshold)
            if pos:
                self.background_click_pro(hwnd, pos[0], pos[1])
                self.log_signal.emit(f"✅ {display} 클릭")
                return True
            self.interruptible_sleep(0.5)
        return False

    def send_text_korean(self, hwnd, text):
        self.check_running()
        for char in text:
            win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            self.interruptible_sleep(0.05)

    def run(self):
        self.is_running = True
        parent_hwnd = win32gui.FindWindow(None, self.GAME_TITLE)
        if not parent_hwnd:
            self.log_signal.emit(f"❌ '{self.GAME_TITLE}' 창을 찾을 수 없습니다.")
            self.finished_signal.emit()
            return
        child_hwnd = win32gui.FindWindowEx(parent_hwnd, None, None, None) or parent_hwnd
        try:
            while self.is_running:
                if self.mode == "fishing":
                    self.run_fishing_logic(parent_hwnd, child_hwnd)
                else:
                    self.run_item_sell_logic(parent_hwnd, child_hwnd)
        except InterruptedError:
            self.log_signal.emit("\n🛑 코드 실행 중지됨.")
        except Exception as e:
            self.log_signal.emit(f"⚠️ 오류 발생: {e}")
        finally:
            self.countdown_signal.emit(-1)
            self.finished_signal.emit()

    def run_fishing_logic(self, parent_hwnd, child_hwnd):
        self.log_signal.emit("\n🔍 낚시 아이콘 확인 중...")
        while self.is_running:
            screen = self.get_window_screenshot(parent_hwnd)
            if self.find_image_pos(screen, 'ready.png'):
                self.log_signal.emit("  ✅ 낚시 준비 완료!")
                break
            self.send_key_perfect(child_hwnd, 0x57, 0x11, 0.01) 
            self.interruptible_sleep(0.3)
        
        self.log_signal.emit("[낚시 캐스팅]...")
        self.send_key_perfect(child_hwnd, 0x57, 0x11, 0.05)
        self.interruptible_sleep(0.3)
        self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
        self.interruptible_sleep(1.0)
        
        start_time = time.time()
        while self.is_running:
            screen = self.get_window_screenshot(parent_hwnd)
            for img in self.fishing_conf['yes']:
                if self.find_image_pos(screen, img, self.fishing_conf['threshold']):
                    self.log_signal.emit("  🎣 물었다!")
                    self.interruptible_sleep(6.0)
                    self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39)
                    self.interruptible_sleep(5.0)
                    self.send_key_perfect(child_hwnd, 0x57, 0x11)
                    self.stats_signal.emit("success")
                    return
            for img in self.fishing_conf['no']:
                if self.find_image_pos(screen, img, self.fishing_conf['threshold']):
                    self.log_signal.emit("  ❌ 놓쳤다!")
                    self.interruptible_sleep(0.5)
                    self.send_key_perfect(child_hwnd, 0x57, 0x11)
                    self.interruptible_sleep(1.0)
                    self.stats_signal.emit("fail")
                    return
            if time.time() - start_time >= self.fishing_conf['timeout']:
                self.log_signal.emit("⚠️ 타임아웃")
                self.send_key_perfect(child_hwnd, 0x57, 0x11)
                return
            self.interruptible_sleep(0.3)

    def run_item_sell_logic(self, parent_hwnd, child_hwnd):
        self.log_signal.emit("\n🔄 매크로 사이클 시작...")
        if self.mode == "magic_stone":
            self.interruptible_sleep(2)
            self.send_key_perfect(child_hwnd, 0x49, 0x17)
            self.interruptible_sleep(1.5)
            self.find_and_click(parent_hwnd, 'item.png', label='아이템 탭')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'ingre.png', label='재료 탭')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'search1.png', label='검색 버튼')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'searchfield.png', label='검색창')
            self.interruptible_sleep(0.8)
            self.log_signal.emit("⌨️ '백동' 입력 중...")
            self.send_text_korean(child_hwnd, "백동")
            self.interruptible_sleep(1)
            self.find_and_click(parent_hwnd, 'apply.png', label='적용')
            self.find_and_click(parent_hwnd, 'backdong.png', label='백동 광석')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'sell.png', label='판매')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'max.png', label='최대 수량')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'selling.png', label='판매 확인')
            self.interruptible_sleep(0.8)
            
            self.find_and_click(parent_hwnd, 'search2.png', label='검색 버튼')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'clearfield.png', label='검색 초기화')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'searchfield.png', label='검색창')
            self.interruptible_sleep(0.8)
            self.log_signal.emit("⌨️ '돌멩이' 입력 중...")
            self.send_text_korean(child_hwnd, "돌멩이")
            self.interruptible_sleep(1)
            self.find_and_click(parent_hwnd, 'apply.png', label='적용')
            self.find_and_click(parent_hwnd, 'dolmeng.png', label='돌멩이')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'sell.png', label='판매')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'max.png', label='최대 수량')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'selling.png', label='판매 확인')
            self.interruptible_sleep(0.8)
            
            self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01)
            self.interruptible_sleep(0.8)
            self.send_key_perfect(child_hwnd, 0x49, 0x17)
            self.interruptible_sleep(1.5)
            self.find_and_click(parent_hwnd, 'ingre.png', label='재료 탭')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'search1.png', label='검색 버튼')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'searchfield.png', label='검색창')
            self.interruptible_sleep(0.8)
            self.log_signal.emit("⌨️ '마력' 입력 중...")
            self.send_text_korean(child_hwnd, "마력")
            self.interruptible_sleep(1)
            self.find_and_click(parent_hwnd, 'apply.png', label='적용')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'madol.png', label='마력돌')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'findmethod.png', label='획득 방법')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'findmethod1.png', label='채집 선택')
            
        else:
            self.send_key_perfect(child_hwnd, 0x49, 0x17)
            self.interruptible_sleep(1.5)
            self.find_and_click(parent_hwnd, 'item.png', label='아이템 탭')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'ingre.png', label='재료 탭')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'search1.png', label='검색 버튼')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'searchfield.png', label='검색창')
            self.interruptible_sleep(0.8)
            self.log_signal.emit("⌨️ '튼튼' 입력 중...")
            self.send_text_korean(child_hwnd, "튼튼")
            self.interruptible_sleep(1)
            self.find_and_click(parent_hwnd, 'apply.png', label='적용')
            self.find_and_click(parent_hwnd, 'tenten.png', label='튼튼버섯')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'sell.png', label='판매')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'max.png', label='최대 수량')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'selling.png', label='판매 확인')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'tentenpo.png', label='튼튼버섯 포자')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'sell.png', label='판매')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'max.png', label='최대 수량')
            self.interruptible_sleep(0.8)
            self.find_and_click(parent_hwnd, 'selling.png', label='판매 확인')
            self.interruptible_sleep(0.8)
            self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01)
            
        self.stats_signal.emit("success")
        wait_next = random.randint(30, 60)
        self.log_signal.emit(f"🏁 사이클 완료. 약 {wait_next}초 후 재시작합니다.")
        for i in range(wait_next, -1, -1):
            self.check_running()
            self.countdown_signal.emit(i)
            time.sleep(1)
        self.countdown_signal.emit(-1)

    def stop(self):
        self.is_running = False


class MainWindow(QWidget):
    """프로그램의 메인 화면(UI) — 모던 디자인"""
    def __init__(self):
        super().__init__()
        self.initUI()
        self.worker = None

    def initUI(self):
        self.setWindowTitle('모비노기 통합 매크로')
        self.setFixedSize(400, 560)
        
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_path, 'my_icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(16, 14, 16, 12)
        self.main_layout.setSpacing(10)
        
        # ── 상단 헤더 ──
        header = QHBoxLayout()
        header.setSpacing(10)
        
        self.title_lbl = QLabel("모비노기 매크로")
        self.title_lbl.setFont(QFont("Malgun Gothic", 13, QFont.Bold))
        
        header.addWidget(self.title_lbl)
        header.addStretch(1)
        self.main_layout.addLayout(header)
        
        # ── 메모 영역 ──
        self.memo_edit = QTextEdit()
        self.memo_edit.setPlaceholderText("메모 입력..")
        self.memo_edit.setFixedHeight(50)
        self.main_layout.addWidget(self.memo_edit)

        # ── 작업 모드 카드 ──
        self.mode_card = QFrame()
        self.mode_card.setObjectName("modeCard")
        mode_card_layout = QVBoxLayout(self.mode_card)
        mode_card_layout.setContentsMargins(14, 10, 14, 10)
        mode_card_layout.setSpacing(8)
        
        mode_title = QLabel("작업 모드")
        mode_title.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        mode_title.setObjectName("cardTitle")
        mode_card_layout.addWidget(mode_title)
        
        radio_layout = QHBoxLayout()
        radio_layout.setSpacing(16)
        self.radio_magic = QRadioButton("마력 깃든 돌")
        self.radio_magic.setChecked(True)
        self.radio_magic.setCursor(Qt.PointingHandCursor)
        self.radio_moss = QRadioButton("반짝이는 이끼")
        self.radio_moss.setCursor(Qt.PointingHandCursor)
        self.radio_fishing = QRadioButton("낚시")
        self.radio_fishing.setCursor(Qt.PointingHandCursor)
        
        radio_layout.addWidget(self.radio_magic)
        radio_layout.addWidget(self.radio_moss)
        radio_layout.addWidget(self.radio_fishing)
        mode_card_layout.addLayout(radio_layout)
        
        self.main_layout.addWidget(self.mode_card)

        # ── 시작 / 중지 버튼 ──
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        
        self.btn_start = QPushButton("▶  시작")
        self.btn_start.setFixedHeight(42)
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        self.btn_start.clicked.connect(self.start_macro)
        
        self.btn_stop = QPushButton("■  중지")
        self.btn_stop.setFixedHeight(42)
        self.btn_stop.setCursor(Qt.PointingHandCursor)
        self.btn_stop.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_macro)
        
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)
        self.main_layout.addLayout(btn_layout)

        # ── 로그 헤더 ──
        log_header = QHBoxLayout()
        self.log_label = QLabel("실행 로그")
        self.log_label.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        self.countdown_lbl = QLabel("")
        self.countdown_lbl.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        self.countdown_lbl.setObjectName("countdown")
        log_header.addWidget(self.log_label)
        log_header.addStretch(1)
        log_header.addWidget(self.countdown_lbl)
        self.main_layout.addLayout(log_header)

        # ── 로그 텍스트 ──
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("logText")
        
        self.btn_clear = QPushButton("지우기", self.log_text)
        self.btn_clear.setFixedWidth(55)
        self.btn_clear.setFixedHeight(22)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self.clear_logs)
        self.main_layout.addWidget(self.log_text)

        # ── 하단 가이드 ──
        guide_layout = QHBoxLayout()
        self.notice_content = QLabel(
            "• 마력 깃든 돌: 얼음협곡 광맥 / 반짝이는 이끼: 시드스넷타 버섯\n"
            "• 낚시: 낚시터 진입 후 앞으로 갔을 때 안 움직이는 자리\n"
            "• 창 크기: 1280x960 설정 후 세로만 최소로 바꾸기"
        )
        self.notice_content.setObjectName("guideText")
        self.notice_content.setFont(QFont("Malgun Gothic", 8))
        
        self.btn_guide = QPushButton("❓")
        self.btn_guide.setFixedSize(31, 31)
        self.btn_guide.setCursor(Qt.PointingHandCursor)
        self.btn_guide.clicked.connect(self.show_guide)
        
        guide_layout.addWidget(self.notice_content)
        guide_layout.addStretch(1)
        guide_layout.addWidget(self.btn_guide, alignment=Qt.AlignTop)
        self.main_layout.addLayout(guide_layout)
        
        self.setLayout(self.main_layout)
        self.apply_style()

    def apply_style(self):
        c = COLORS
        
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {c['bg']};
                color: {c['text']};
                font-family: 'Malgun Gothic';
            }}
            
            /* 메모 입력 */
            QTextEdit#logText {{
                background-color: {c['log_bg']};
                color: {c['log_text']};
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                padding: 6px;
            }}
            QTextEdit {{
                background-color: {c['input_bg']};
                color: {c['text']};
                border: 1px solid {c['input_border']};
                border-radius: 8px;
                padding: 6px;
                font-size: 10pt;
            }}
            
            /* 모드 카드 */
            QFrame#modeCard {{
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 10px;
            }}
            QLabel#cardTitle {{
                color: {c['accent']};
                border: none;
                background: transparent;
            }}
            
            /* 라디오 */
            QRadioButton {{
                color: {c['text']};
                spacing: 6px;
                font-size: 10pt;
                background: transparent;
                border: none;
            }}
            QRadioButton::indicator {{
                width: 16px;
                height: 16px;
                border-radius: 8px;
                border: 2px solid {c['text_dim']};
                background-color: transparent;
            }}
            QRadioButton::indicator:checked {{
                border: 2px solid {c['accent']};
                background-color: {c['accent']};
            }}
            QRadioButton::indicator:hover {{
                border: 2px solid {c['accent_hover']};
            }}
            
            /* 시작 버튼 */
            QPushButton#btnStart {{
                background-color: {c['success']};
                color: white;
                border: none;
                border-radius: 10px;
                font-weight: bold;
            }}
            QPushButton#btnStart:hover {{
                background-color: {c['success_hover']};
            }}
            QPushButton#btnStart:pressed {{
                background-color: #009978;
            }}
            QPushButton#btnStart:disabled {{
                background-color: {c['card']};
                color: {c['text_dim']};
                border: 1px solid {c['card_border']};
            }}
            
            /* 중지 버튼 */
            QPushButton#btnStop {{
                background-color: {c['danger']};
                color: white;
                border: none;
                border-radius: 10px;
                font-weight: bold;
            }}
            QPushButton#btnStop:hover {{
                background-color: {c['danger_hover']};
            }}
            QPushButton#btnStop:pressed {{
                background-color: #c0392b;
            }}
            QPushButton#btnStop:disabled {{
                background-color: {c['card']};
                color: {c['text_dim']};
                border: 1px solid {c['card_border']};
            }}
            
            QPushButton#btnTheme {{
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 16px;
                font-size: 14px;
            }}
            QPushButton#btnTheme:hover {{
                background-color: {c['surface']};
                border: 1px solid {c['accent']};
            }}
            
            /* 가이드 버튼 */
            QPushButton#btnGuide {{
                background-color: transparent;
                border: 1px solid {c['card_border']};
                border-radius: 12px;
                font-size: 12px;
            }}
            QPushButton#btnGuide:hover {{
                border: 1px solid {c['accent']};
            }}
            
            /* 지우기 버튼 */
            QPushButton#btnClear {{
                background-color: {c['card']};
                color: {c['text_dim']};
                border: 1px solid {c['card_border']};
                border-radius: 4px;
                font-size: 9pt;
            }}
            QPushButton#btnClear:hover {{
                color: {c['text']};
                border: 1px solid {c['accent']};
            }}
            
            /* 카운트다운 */
            QLabel#countdown {{
                color: {c['danger']};
                background: transparent;
                border: none;
            }}
            
            /* 가이드 텍스트 */
            QLabel#guideText {{
                color: {c['text_dim']};
                background: transparent;
                border: none;
            }}
            
            /* 스크롤바 */
            QScrollBar:vertical {{
                background-color: transparent;
                width: 8px;
                margin: 0;
            }}
            QScrollBar::handle:vertical {{
                background-color: {c['card_border']};
                border-radius: 4px;
                min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{
                background-color: {c['accent']};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
            }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: transparent;
            }}
        """)
        
        # objectName 설정 (스타일시트 적용 후 재설정)
        self.btn_start.setObjectName("btnStart")
        self.btn_stop.setObjectName("btnStop")
        self.btn_guide.setObjectName("btnGuide")
        self.btn_clear.setObjectName("btnClear")
        self.log_text.setObjectName("logText")
        
        # objectName 변경 후 스타일 강제 새로고침
        for w in [self.btn_start, self.btn_stop, self.btn_guide, self.btn_clear, self.log_text]:
            w.style().unpolish(w)
            w.style().polish(w)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'btn_clear'):
            self.btn_clear.move(self.log_text.width() - 65, 5)

    def start_macro(self):
        mode = "magic_stone" if self.radio_magic.isChecked() else "moss" if self.radio_moss.isChecked() else "fishing"
        self.worker = MacroWorker(mode)
        self.worker.log_signal.connect(self.append_log)
        self.worker.countdown_signal.connect(self.update_countdown)
        self.worker.finished_signal.connect(self.on_finished)
        
        self.thread = threading.Thread(target=self.worker.run)
        self.thread.daemon = True
        self.thread.start()
        
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        mode_name = "마력 깃든 돌" if mode == "magic_stone" else "반짝이는 이끼" if mode == "moss" else "낚시"
        self.append_log(f"=== {mode_name} 매크로 시작 ===")

    def stop_macro(self):
        if self.worker:
            self.worker.stop()
            self.btn_stop.setEnabled(False)

    def clear_logs(self):
        self.log_text.clear()
        self.append_log("── 로그 초기화 ──")

    def on_finished(self):
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.countdown_lbl.setText("")

    def show_guide(self):
        notice = NoticeDialog(self)
        notice.exec_()

    @pyqtSlot(str)
    def append_log(self, text):
        self.log_text.append(text)
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    @pyqtSlot(int)
    def update_countdown(self, seconds):
        if seconds >= 0:
            self.countdown_lbl.setText(f"다음 실행까지: {seconds}초")
        else:
            self.countdown_lbl.setText("")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    
    notice = NoticeDialog(ex)
    notice.exec_()
    sys.exit(app.exec_())
