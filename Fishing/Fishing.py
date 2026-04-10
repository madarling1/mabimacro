import sys
import os
import time
import threading
import cv2
import numpy as np
import ctypes
import win32gui
import win32ui
import win32con
import win32api
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTextEdit, QLabel,
                             QDialog, QTextBrowser)
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, Qt
from PyQt5.QtGui import QFont, QIcon

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


# 컬러 팔레트
LIGHT_COLORS = {
    'bg':           '#e7f2eb',
    'card':         '#f8fcf9',
    'card_border':  '#5f8f73',
    'surface':      '#eef7f1',
    'text':         '#1f3328',
    'text_dim':     '#5c7768',
    'accent':       '#2f9e66',
    'accent_hover': '#45b87c',
    'success':      '#34b36f',
    'success_hover':'#4cc786',
    'danger':       '#e07a7a',
    'danger_hover': '#ec9292',
    'log_bg':       '#f4faf6',
    'log_text':     '#2a5a43',
    'input_bg':     '#f9fdfb',
    'input_border': '#5f8f73',
}

DARK_COLORS = {
    'bg':           '#121b15',
    'card':         '#1b2720',
    'card_border':  '#3e6a52',
    'surface':      '#1f2d25',
    'text':         '#e4f2e8',
    'text_dim':     '#9fbea9',
    'accent':       '#53c083',
    'accent_hover': '#69cf95',
    'success':      '#42b978',
    'success_hover':'#57c98a',
    'danger':       '#d67979',
    'danger_hover': '#e08d8d',
    'log_bg':       '#142019',
    'log_text':     '#8ce0b2',
    'input_bg':     '#16231b',
    'input_border': '#3e6a52',
}

COLORS = DARK_COLORS.copy()

WINDOW_TITLE = "마비노기 모바일"
TARGET_CLIENT_WIDTH = 1280
TARGET_CLIENT_HEIGHT = 691


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class NoticeDialog(QDialog):
    """매크로 시작 전, 사용자가 꼭 알아야 할 주의사항과 가이드를 보여주는 안내창입니다."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI(parent)

    def initUI(self, parent):
        self.setWindowTitle('모비노기 낚시 매크로 가이드')
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
        <p>이 프로그램은 <b>집중해서 낚시하기 전용</b>입니다. <br>준비 아이콘을 감지 후 캐스팅하고, 성공 또는 실패를 이미지로 판별해 낚시를 진행합니다.</p>
        <hr>
        <p><b>1. 시작 위치</b></p>
        <ul>
            <li>낚시 : 원하는 낚시터 진입 후 앞으로 가도 안 움직이는 자리</li>
        </ul>
        <p><b>2. 매크로 실행 전 필수 마비노기 세팅</b></p>
        <ul>
            <li>상단의 <b>창 크기 맞추기</b> 버튼으로 게임 창 크기를 맞춰주세요.</li>
            <li>낚시터 입장 후 <b>캐릭터가 더 이상 앞으로 가지 않는 자리</b>에서 시작하세요.</li>
            <li>이미지 감지 방식이라 <b>게임 창이 최소화 되면 안 됩니다.</b></li>
        </ul>
        <p><b>3. 주의사항</b></p>
        <ul>
            <li>펫 상호작용 때문에 오류가 잦으니 <b>펫은 꼭 넣고 진행</b>해주세요.</li>
            <li>게임 창 크기가 변하면 오류가 발생할 수 있습니다. 창 크기를 맞춰주세요</li>
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


class TimeoutRecoveryError(Exception):
    """로그 기반 타임아웃 발생 시 raise"""


class MacroWorker(QObject):
    """실제로 게임 화면을 분석하고 조작하는 낚시 전용 워커입니다."""

    log_signal = pyqtSignal(str)
    countdown_signal = pyqtSignal(int)
    finished_signal = pyqtSignal()
    stats_signal = pyqtSignal(str)

    LOG_TIMEOUT = 30
    RECOVERY_MAX_ESC = 10

    def __init__(self):
        super().__init__()
        self.is_running = False
        self.GAME_TITLE = WINDOW_TITLE
        self.last_log_time = time.time()
        self.timeout_check_enabled = False
        self.fishing_conf = {
            'ready': ['ready.png'],
            'yes': ['yes1.png', 'yes2.png'],
            'no': ['no1.png', 'no2.png'],
            'threshold': 0.8,
            'timeout': 10.0
        }

    def log(self, msg):
        self.last_log_time = time.time()
        self.log_signal.emit(msg)

    def check_running(self):
        if not self.is_running:
            raise InterruptedError("중지됨")

    def interruptible_sleep(self, seconds):
        steps = int(seconds / 0.1)
        for _ in range(steps):
            self.check_running()
            if self.timeout_check_enabled and (time.time() - self.last_log_time > self.LOG_TIMEOUT):
                raise TimeoutRecoveryError(f"로그 {self.LOG_TIMEOUT}초 미출력 타임아웃")
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

    def send_key_perfect(self, hwnd, vk_code, scan_code, duration=0.1):
        self.check_running()
        win32gui.PostMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        down_lparam = 1 | (scan_code << 16)
        up_lparam = 1 | (scan_code << 16) | (1 << 30) | (1 << 31)
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, down_lparam)
        self.interruptible_sleep(duration)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, up_lparam)

    def find_image_pos(self, screen, image_name, threshold=0.7):
        if screen is None:
            return None

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        full_path = os.path.join(base_path, image_name)
        if not os.path.exists(full_path):
            return None

        try:
            img_array = np.fromfile(full_path, np.uint8)
            template = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            if template is None:
                return None

            res = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(res)
            if max_val >= threshold:
                th, tw = template.shape[:2]
                return (max_loc[0] + tw // 2, max_loc[1] + th // 2)
        except:
            pass

        return None

    def recovery_reset(self, parent_hwnd, child_hwnd):
        self.log("═══ 에러 탈출 ═══")
        for _ in range(self.RECOVERY_MAX_ESC):
            self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
            self.log("  Esc ")
            self.interruptible_sleep(1.5)

            screen = self.get_window_screenshot(parent_hwnd)
            if self.find_image_pos(screen, 'Process.png', threshold=0.7):
                self.log("  메인창 감지 — 복구 완료")
                self.log("═══ 에러 탈출 종료 ═══")
                return True

        self.log("⚠️ 메인창 미감지 — 복구 실패 (최대 Esc 횟수 초과)")
        self.log("🔧 ═══ 에러 탈출 종료 ═══")
        return False

    def run(self):
        self.is_running = True
        parent_hwnd = win32gui.FindWindow(None, self.GAME_TITLE)
        if not parent_hwnd:
            self.log(f"❌ '{self.GAME_TITLE}' 창을 찾을 수 없습니다.")
            self.finished_signal.emit()
            return

        child_hwnd = win32gui.FindWindowEx(parent_hwnd, None, None, None) or parent_hwnd
        try:
            while self.is_running:
                try:
                    self.timeout_check_enabled = True
                    self.run_fishing_logic(parent_hwnd, child_hwnd)
                except TimeoutRecoveryError as e:
                    self.log(f"⏰ 타임아웃: {e}")
                    self.timeout_check_enabled = False
                    if self.recovery_reset(parent_hwnd, child_hwnd):
                        self.log("🔄 복구 완료 — 사이클 재시작")
                        continue
                    self.log("❌ 복구 실패 → 매크로 중지")
                    break
                finally:
                    self.timeout_check_enabled = False
        except InterruptedError:
            self.log("\n🛑 코드 실행 중지됨.")
        except Exception as e:
            self.log(f"⚠️ 오류 발생: {e}")
        finally:
            self.countdown_signal.emit(-1)
            self.finished_signal.emit()

    def run_fishing_logic(self, parent_hwnd, child_hwnd):
        self.log("\n🔍 낚시 아이콘 확인 중...")
        while self.is_running:
            screen = self.get_window_screenshot(parent_hwnd)
            if self.find_image_pos(screen, 'ready.png'):
                self.log("  ✅ 낚시 준비 완료!")
                break
            self.send_key_perfect(child_hwnd, 0x57, 0x11, 0.01)
            self.interruptible_sleep(0.3)

        self.log("[낚시 캐스팅]...")
        self.send_key_perfect(child_hwnd, 0x57, 0x11, 0.05)
        self.interruptible_sleep(0.3)
        self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
        self.interruptible_sleep(1.0)

        start_time = time.time()
        while self.is_running:
            screen = self.get_window_screenshot(parent_hwnd)
            for img in self.fishing_conf['yes']:
                if self.find_image_pos(screen, img, self.fishing_conf['threshold']):
                    self.log("  🎣 물었다!")
                    self.interruptible_sleep(6.0)
                    self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39)
                    self.interruptible_sleep(5.0)
                    self.send_key_perfect(child_hwnd, 0x57, 0x11)
                    self.stats_signal.emit("success")
                    return

            for img in self.fishing_conf['no']:
                if self.find_image_pos(screen, img, self.fishing_conf['threshold']):
                    self.log("  ❌ 놓쳤다!")
                    self.interruptible_sleep(0.5)
                    self.send_key_perfect(child_hwnd, 0x57, 0x11)
                    self.interruptible_sleep(1.0)
                    self.stats_signal.emit("fail")
                    return

            if time.time() - start_time >= self.fishing_conf['timeout']:
                self.log("⚠️ 타임아웃")
                self.send_key_perfect(child_hwnd, 0x57, 0x11)
                return

            self.interruptible_sleep(0.3)

    def stop(self):
        self.is_running = False


class MainWindow(QWidget):
    """프로그램의 메인 화면(UI) — 낚시 전용으로 재구성."""

    def __init__(self):
        super().__init__()
        self.colors = DARK_COLORS.copy()
        self.initUI()
        self.worker = None

    def initUI(self):
        self.setWindowTitle('모비노기 낚시 매크로')
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

        header = QHBoxLayout()
        header.setSpacing(10)

        self.title_lbl = QLabel("모비노기 낚시 매크로")
        self.title_lbl.setFont(QFont("Malgun Gothic", 13, QFont.Bold))

        header.addWidget(self.title_lbl)
        header.addStretch(1)

        self.btn_resize = QPushButton("창 크기 맞추기")
        self.btn_resize.setObjectName("btnResize")
        self.btn_resize.setFixedHeight(30)
        self.btn_resize.setCursor(Qt.PointingHandCursor)
        self.btn_resize.setFont(QFont("Malgun Gothic", 8))
        self.btn_resize.setToolTip("게임 창 크기 맞추기 (1280x691)")
        self.btn_resize.clicked.connect(self.resize_game_window)
        header.addWidget(self.btn_resize)

        self.main_layout.addLayout(header)

        self.memo_edit = QTextEdit()
        self.memo_edit.setPlaceholderText("낚시 메모 입력..")
        self.memo_edit.setFixedHeight(72)
        self.main_layout.addWidget(self.memo_edit)

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

        log_header = QHBoxLayout()
        self.log_label = QLabel("낚시 로그")
        self.log_label.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        self.countdown_lbl = QLabel("")
        self.countdown_lbl.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        self.countdown_lbl.setObjectName("countdown")
        log_header.addWidget(self.log_label)
        log_header.addStretch(1)
        log_header.addWidget(self.countdown_lbl)
        self.main_layout.addLayout(log_header)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("logText")

        self.btn_clear = QPushButton("지우기", self.log_text)
        self.btn_clear.setFixedWidth(55)
        self.btn_clear.setFixedHeight(22)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self.clear_logs)
        self.main_layout.addWidget(self.log_text)

        guide_layout = QHBoxLayout()
        self.notice_content = QLabel(
            "• 자리: 낚시터 진입 후 앞으로 가도 안 움직이는 자리\n"
            "• 준비: 펫 넣기 / 게임 창 최소화 되지 않게 유지\n"
            "• 창 크기: 우상단 버튼으로 게임창 크기 맞추기"
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
        c = self.colors

        self.setStyleSheet(f"""
            QWidget {{
                background-color: {c['bg']};
                color: {c['text']};
                font-family: 'Malgun Gothic';
            }}

            QTextEdit#logText {{
                background-color: {c['log_bg']};
                color: {c['log_text']};
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                border: 1px solid {c['card_border']};
                border-radius: 10px;
                padding: 8px;
            }}
            QTextEdit {{
                background-color: {c['input_bg']};
                color: {c['text']};
                border: 1px solid {c['input_border']};
                border-radius: 10px;
                padding: 8px;
                font-size: 10pt;
                selection-background-color: {c['accent']};
                selection-color: white;
            }}

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

            QPushButton#btnResize {{
                background-color: {c['card']};
                color: {c['text']};
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                padding: 0 10px;
            }}
            QPushButton#btnResize:hover {{
                background-color: {c['surface']};
                border: 1px solid {c['accent']};
            }}

            QPushButton#btnGuide {{
                background-color: transparent;
                border: 1px solid {c['card_border']};
                border-radius: 12px;
                font-size: 12px;
            }}
            QPushButton#btnGuide:hover {{
                border: 1px solid {c['accent']};
            }}

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

            QLabel#countdown {{
                color: {c['danger']};
                background: transparent;
                border: none;
            }}

            QLabel#guideText {{
                color: {c['text_dim']};
                background: transparent;
                border: none;
            }}

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

        self.btn_start.setObjectName("btnStart")
        self.btn_stop.setObjectName("btnStop")
        self.btn_resize.setObjectName("btnResize")
        self.btn_guide.setObjectName("btnGuide")
        self.btn_clear.setObjectName("btnClear")
        self.log_text.setObjectName("logText")

        for w in [self.btn_start, self.btn_stop, self.btn_resize, self.btn_guide, self.btn_clear, self.log_text]:
            w.style().unpolish(w)
            w.style().polish(w)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'btn_clear'):
            self.btn_clear.move(self.log_text.width() - 65, 5)

    def resize_game_window(self):
        hwnd = win32gui.FindWindow(None, WINDOW_TITLE)
        if not hwnd:
            self.append_log(f"❌ '{WINDOW_TITLE}' 창을 찾지 못했습니다.")
            return

        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        has_menu = bool(win32gui.GetMenu(hwnd))

        rect = RECT(0, 0, TARGET_CLIENT_WIDTH, TARGET_CLIENT_HEIGHT)
        ctypes.windll.user32.AdjustWindowRectEx(
            ctypes.byref(rect),
            style,
            has_menu,
            ex_style
        )

        outer_width = rect.right - rect.left
        outer_height = rect.bottom - rect.top
        win32gui.MoveWindow(hwnd, left, top, outer_width, outer_height, True)

        client_left, client_top, client_right, client_bottom = win32gui.GetClientRect(hwnd)
        client_width = client_right - client_left
        client_height = client_bottom - client_top
        self.append_log(f"🪟 게임 창 크기 맞춤: {client_width} x {client_height}")

    def start_macro(self):
        self.worker = MacroWorker()
        self.worker.log_signal.connect(self.append_log)
        self.worker.countdown_signal.connect(self.update_countdown)
        self.worker.finished_signal.connect(self.on_finished)

        self.thread = threading.Thread(target=self.worker.run)
        self.thread.daemon = True
        self.thread.start()

        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.append_log("=== 낚시 매크로 시작 ===")

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
