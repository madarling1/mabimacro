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
                             QPushButton, QTextEdit, QLabel, QCheckBox, QGroupBox, 
                             QDialog, QTextBrowser, QGridLayout, QFrame,
                             QGraphicsDropShadowEffect)
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, Qt
from PyQt5.QtGui import (QFont, QIcon, QColor)

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

COLORS = LIGHT_COLORS.copy()

# ── 카테고리별 작업 모드 정의 ──
CATEGORIES = [
    {
        "name": "금속",
        "processing": "MetalProcessing.png",
        "items": [
            ("철괴", "IronIngot.png"),
            ("강철괴", "SteelIngot.png"),
            ("합금강괴", "CopperIngot.png"),
        ]
    },
    {
        "name": "목재",
        "processing": "WoodProcessing.png",
        "items": [
            ("목재", "Wood.png"),
            ("목재+", "WoodPlus.png"),
            ("상급목재", "FineWood.png"),
        ]
    },
    {
        "name": "가죽",
        "processing": "LeatherProcessing.png",
        "items": [
            ("가죽", "Leather.png"),
            ("가죽+", "LeatherPlus.png"),
            ("상급가죽", "FineLeather.png"),
        ]
    },
    {
        "name": "옷감",
        "processing": "FabricProcessing.png",
        "items": [
            ("옷감", "Fabric.png"),
            ("옷감+", "FabricPlus.png"),
            ("상급옷감", "FineFabric.png"),
        ]
    },
]


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
        <h3 style='color: #6c5ce7;'>[ 필독: 가공 매크로 가이드 ]</h3>
        <p>이 매크로는 마비노기 모바일의 가공 자동화를 위해 제작되었습니다.</p>
        <p>선택한 재료를 자동으로 인식하여 자동으로 가공합니다.</p>
        <p><b>여러 재료를 동시에 선택하면 자동으로 제작대를 전환하며 로테이션합니다.</b></p>
        <p><b>제작 재료는 자동화 되어있지 않으니, 재료를 구비한 후 실행해주세요.</b></p>
        <hr>
        <p><b>1. 지원 재료</b></p>
        <ul>
            <li>금속: 철괴, 강철괴, 합금강괴</li>
            <li>목재: 목재, 목재+, 상급목재</li>
            <li>가죽: 가죽, 가죽+, 상급가죽</li>
            <li>옷감: 옷감, 옷감+, 상급옷감</li>
        </ul>
        <p><b>2. 매크로 실행 전 필수 세팅</b></p>
        <ul>
            <li>창 크기: <b>1280x960</b> 설정 후, <b>세로만 최소</b>로 줄여주세요.</li>
            <li><b>제작대에 가서 가공 창을 열어놓은 상태</b>에서 매크로를 시작하세요.</li>
            <li>복수 선택 시: 반드시 <b>선택한 재료 중 아무 제작대나 열어놓고</b> 시작하세요. (자동 감지)</li>
        </ul>
        <p><b>3. 주의사항</b></p>
        <ul>
            <li>이미지 스캔 방식이라 창이 가려지면 안 됩니다.</li>
            <li>각 분야에서 하나씩만 선택 가능합니다 (최대 4개 동시).</li>
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
    pass


class MacroWorker(QObject):
    """통합 제작 매크로 워커 — 단일/복수 재료 로테이션 지원"""
    log_signal = pyqtSignal(str)
    countdown_signal = pyqtSignal(int)
    finished_signal = pyqtSignal()

    GAME_TITLE = "마비노기 모바일"
    VK_K = 0x4B  # K key

    # 타임아웃 상수
    LOG_TIMEOUT = 30       # 로그 미출력 최대 허용(초)
    RECOVERY_MAX_ESC = 10  # 복구 시 Esc 최대 횟수

    def __init__(self, targets):
        """
        targets: list of dict
        각 항목: { "label": "철괴", "image": "IronIngot.png", 
                   "category": "금속", "processing": "MetalProcessing.png" }
        """
        super().__init__()
        self.is_running = False
        self.original_targets = list(targets)  # 초기 선택 목록
        self.active_targets = list(targets)    # 현재 활성 목록 (Error시 제거)
        self.threshold = 0.9
        self.last_log_time = time.time()
        self.timeout_check_enabled = False  # run_transition 내에서만 활성화

        if getattr(sys, 'frozen', False):
            self.base_path = sys._MEIPASS
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))

    def log(self, msg):
        """로그 출력 + 타임아웃 타이머 리셋"""
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
        if screen is None:
            return None
        full_path = os.path.join(self.base_path, image_name)
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

    def find_and_click(self, hwnd, image_name, threshold=0.7, wait_time=5):
        """이미지 감지 후 클릭. 성공 시 True, 타임아웃 시 False."""
        start_time = time.time()
        while time.time() - start_time < wait_time:
            self.check_running()
            screen = self.get_window_screenshot(hwnd)
            pos = self.find_image_pos(screen, image_name, threshold)
            if pos:
                self.background_click_pro(hwnd, pos[0], pos[1])
                return True
            self.interruptible_sleep(0.5)
        return False

    def recovery_reset(self, parent_hwnd, child_hwnd):
        """타임아웃 복구: Esc 반복 → Process.png 감지될 때까지"""
        self.log("🔧 ═══ 타임아웃 복구 시작 ═══")
        for attempt in range(self.RECOVERY_MAX_ESC):
            self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
            self.log(f"⌨️ Esc ({attempt + 1}/{self.RECOVERY_MAX_ESC}회)")
            self.interruptible_sleep(1.5)

            screen = self.get_window_screenshot(parent_hwnd)
            if self.find_image_pos(screen, 'Process.png', threshold=0.7):
                self.log("✅ Process.png 감지 — 복구 완료!")
                self.log("🔧 ═══ 타임아웃 복구 종료 ═══")
                return True

        self.log("⚠️ Process.png 미감지 — 복구 실패 (최대 Esc 횟수 초과)")
        self.log("🔧 ═══ 타임아웃 복구 종료 ═══")
        return False

    def check_error(self, hwnd):
        """Error 이미지 감지. 반환값: None / 'NoIngre' / 'FullOrder'"""
        screen = self.get_window_screenshot(hwnd)
        if self.find_image_pos(screen, 'Error1_NoIngre.png', threshold=self.threshold):
            self.log("⚠️ 재료 부족 감지")
            return 'NoIngre'
        if self.find_image_pos(screen, 'Error2_FullOrder.png', threshold=self.threshold):
            self.log("⚠️ 주문 가득 참 감지")
            return 'FullOrder'
        return None

    # ═══════════════════════════════════════════════
    #  메인 실행
    # ═══════════════════════════════════════════════
    def run(self):
        self.is_running = True
        parent_hwnd = win32gui.FindWindow(None, self.GAME_TITLE)
        if not parent_hwnd:
            self.log(f"❌ '{self.GAME_TITLE}' 창을 찾을 수 없습니다.")
            self.finished_signal.emit()
            return
        child_hwnd = win32gui.FindWindowEx(parent_hwnd, None, None, None) or parent_hwnd

        try:
            if len(self.active_targets) == 1:
                self.run_single(parent_hwnd, child_hwnd)
            else:
                self.run_rotation(parent_hwnd, child_hwnd)
        except InterruptedError:
            self.log("\n🛑 매크로 중지됨.")
        except Exception as e:
            self.log(f"⚠️ 오류 발생: {e}")
        finally:
            self.countdown_signal.emit(-1)
            self.finished_signal.emit()

    # ═══════════════════════════════════════════════
    #  단일 선택 모드 (기존 로직)
    # ═══════════════════════════════════════════════
    def run_single(self, parent_hwnd, child_hwnd):
        """단일 재료 — 기존 배포용2.py와 동일한 로직"""
        target = self.active_targets[0]
        img = target["image"]
        self.log(f"\n📌 단일 모드: {target['label']}")

        while self.is_running:
            self.log(f"\n🔄 [{target['label']}] 제작 사이클 시작...")

            # 1. 타겟 찾기 → 클릭 → 스페이스바
            screen = self.get_window_screenshot(parent_hwnd)
            pos = self.find_image_pos(screen, img, threshold=self.threshold)
            if pos:
                self.background_click_pro(parent_hwnd, pos[0], pos[1])
                self.log(f"⛏️ {target['label']} 클릭")
                self.interruptible_sleep(1.5)
                self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                self.log("⌨️ 가공 시작")
                self.interruptible_sleep(1.5)

            # 2. Error 체크
            if self.check_error(parent_hwnd):
                self.log("⚠️ 에러 감지 → 수거 진행")
                self.interruptible_sleep(1.5)
                self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
                self.log("⌨️ 창 닫기")
                self.interruptible_sleep(1.5)

                # ReceiveAll 대기
                self.log("⏳ 모두 받기 버튼 대기 중...")
                while self.is_running:
                    screen = self.get_window_screenshot(parent_hwnd)
                    pos = self.find_image_pos(screen, 'ReceiveAll.png', threshold=self.threshold)
                    if pos:
                        self.background_click_pro(parent_hwnd, pos[0], pos[1])
                        self.log("✅ 모두 받기 완료")
                        break
                    self.interruptible_sleep(0.5)
                self.check_running()
                self.interruptible_sleep(1.5)
                self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                self.log("⌨️ 확인")
                self.interruptible_sleep(1.5)

                # 7-2. 10초 대기 → ReceiveAll 재감지 → 스페이스바
                self.log("⏳ 10초 대기 (제작대 이동)...")
                self.interruptible_sleep(10.0)
                self.log("🔍 모두 받기 버튼 재탐색 중...")
                while self.is_running:
                    screen = self.get_window_screenshot(parent_hwnd)
                    pos = self.find_image_pos(screen, 'ReceiveAll.png', threshold=self.threshold)
                    if pos:
                        self.background_click_pro(parent_hwnd, pos[0], pos[1])
                        self.log("✅ 모두 받기 완료")
                        break
                    self.interruptible_sleep(1.0)
                self.check_running()
                self.interruptible_sleep(1.5)
                self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                self.log("⌨️ 확인")
                self.interruptible_sleep(1.5)

    # ═══════════════════════════════════════════════
    #  복수 선택 — 로테이션 모드
    # ═══════════════════════════════════════════════
    def detect_current_station(self, parent_hwnd):
        """현재 열려있는 제작대가 어떤 재료인지 자동 감지"""
        screen = self.get_window_screenshot(parent_hwnd)
        for idx, target in enumerate(self.active_targets):
            pos = self.find_image_pos(screen, target["image"], threshold=self.threshold)
            if pos:
                return idx
        return 0  # 감지 실패 시 첫 번째로 시작

    def run_rotation(self, parent_hwnd, child_hwnd):
        """복수 재료 로테이션 — 제작대 자동 전환"""
        names = ", ".join([t["label"] for t in self.active_targets])
        self.log(f"\n📌 로테이션 모드: {names}")

        # 현재 열려있는 제작대 자동 감지
        current_idx = self.detect_current_station(parent_hwnd)
        detected = self.active_targets[current_idx]
        self.log(f"🔍 현재 제작대 감지: {detected['label']}")

        while self.is_running and len(self.active_targets) > 0:
            # 1개만 남으면 해당 제작대로 이동 후 단일 모드로 전환
            if len(self.active_targets) == 1:
                self.log("\n📌 1개만 남아 단일 모드로 전환")
                last = self.active_targets[0]
                # Processing1/2 → [분야]Processing 으로 이동
                self.log("🔍 가공 메뉴 탐색 중...")
                for proc_img in ['Processing1.png', 'Processing2.png']:
                    if self.find_and_click(parent_hwnd, proc_img, threshold=0.7, wait_time=3):
                        self.log("✅ 가공 메뉴 클릭")
                        break
                self.interruptible_sleep(1.5)
                self.log(f"🔍 {last['category']} 가공대 이동 중...")
                if self.find_and_click(parent_hwnd, last['processing'], threshold=0.7, wait_time=5):
                    self.log(f"✅ {last['category']} 가공대 클릭")
                self.log("⏳ 10초 대기...")
                self.interruptible_sleep(10.0)
                # ReceiveAll 3초 감지
                self.log("🔍 모두 받기 버튼 3초 탐색 중...")
                if self.find_and_click(parent_hwnd, 'ReceiveAll.png', threshold=self.threshold, wait_time=3):
                    self.log("✅ 모두 받기 완료")
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                    self.log("⌨️ 확인")
                    self.interruptible_sleep(1.5)
                else:
                    self.log("⏩ 모두 받기 없음, 단일 모드 진행")
                self.run_single(parent_hwnd, child_hwnd)
                return

            current_idx = current_idx % len(self.active_targets)
            current = self.active_targets[current_idx]

            # ═══ 제작 루프 ═══
            self.run_craft_loop(parent_hwnd, child_hwnd, current)

            # ═══ 중간 파트: 다음 제작대로 이동 (FullOrder 시 연쇄 전환) ═══
            next_idx = (current_idx + 1) % len(self.active_targets)
            skip_to_step5 = False  # 첫 전환은 step 4부터
            skip_k = False

            while self.is_running and len(self.active_targets) > 1:
                next_target = self.active_targets[next_idx % len(self.active_targets)]
                result = self.run_transition(parent_hwnd, child_hwnd, next_target, skip_to_step5=skip_to_step5, skip_k=skip_k)

                if result == "ok":
                    current_idx = next_idx % len(self.active_targets)
                    break
                elif result == "error":
                    # NoIngre → 제외, Esc 2회, step 4부터 다음 타겟
                    self.log(f"🚫 {next_target['label']} 로테이션에서 제외")
                    self.active_targets.remove(next_target)
                    if len(self.active_targets) == 0:
                        self.log("❌ 모든 재료 소진. 매크로 종료.")
                        return
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
                    self.log("⌨️ 에러 창 닫기 (1/2)")
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
                    self.log("⌨️ 에러 창 닫기 (2/2)")
                    self.interruptible_sleep(1.5)
                    # K 없이 Processing1/2부터 다음 타겟으로
                    skip_to_step5 = False
                    skip_k = True
                    continue
                elif result == "fullorder":
                    # FullOrder → Esc 2회 완료, 다음 타겟의 step 5로
                    next_idx = (next_idx + 1) % len(self.active_targets)
                    skip_to_step5 = True
                    self.log(f"🔁 다음 재료 제작대로 전환...")
                    continue

    def run_craft_loop(self, parent_hwnd, child_hwnd, target):
        """제작 루프: ReceiveAll 감지 → 재료 클릭 반복 → Error시 Esc → Esc(닫기)"""
        img = target["image"]
        label = target["label"]
        self.log(f"\n🔄 [{label}] 제작 루프 시작...")

        try:
            self.timeout_check_enabled = True

            # --- 1. ReceiveAll.png 3초 감지 ---
            self.log("🔍 모두 받기 버튼 3초 탐색 중...")
            found = self.find_and_click(parent_hwnd, 'ReceiveAll.png', threshold=self.threshold, wait_time=3)
            if found:
                self.log("✅ 모두 받기 완료")
                self.interruptible_sleep(1.5)
                self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                self.log("⌨️ 확인")
                self.interruptible_sleep(1.5)
            else:
                self.log("⏩ 모두 받기 없음, 제작 진행") 

            # --- 2. 재료 클릭 → 스페이스바 반복, Error시 탈출 ---
            while self.is_running:
                screen = self.get_window_screenshot(parent_hwnd)
                pos = self.find_image_pos(screen, img, threshold=self.threshold)
                if pos:
                    self.background_click_pro(parent_hwnd, pos[0], pos[1])
                    self.log(f"⛏️ {label} 클릭")
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                    self.log("⌨️ 제작 시작")
                    self.interruptible_sleep(1.5)
                else:
                    self.interruptible_sleep(0.5)

                # Error 체크
                if self.check_error(parent_hwnd):
                    self.log("⚠️ 에러 감지 → 제작 중단")
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
                    self.interruptible_sleep(1.5)
                    break

            # --- 3. Esc (제작대 닫기) ---
            self.interruptible_sleep(1.5)
            self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
            self.log("⌨️ 제작대 닫기")   
            self.interruptible_sleep(1.5)

        except TimeoutRecoveryError as e:
            self.log(f"⏰ 제작 루프 타임아웃: {e}")
            self.timeout_check_enabled = False
            if self.recovery_reset(parent_hwnd, child_hwnd):
                self.log("🔄 복구 완료 — 다음 전환 단계로 진행")
            else:
                self.log("❌ 복구 실패 → 매크로 중지")
                self.is_running = False
        finally:
            self.timeout_check_enabled = False

    def run_transition(self, parent_hwnd, child_hwnd, next_target, skip_to_step5=False, skip_k=False):
        """중간 파트: 다음 제작대로 이동 및 초기 제작 시도.
        반환값: "ok" / "error" / "fullorder"
        skip_to_step5=True: step 4 전체 건너뛰고 step 5부터 (FullOrder 후)
        skip_k=True: K 입력만 건너뛰고 Processing1/2부터 (NoIngre 제외 후)

        로그 기반 타임아웃: 30초간 로그 미출력 시 recovery_reset → Step 4(K키)부터 재시작
        """
        next_label = next_target["label"]
        next_img = next_target["image"]
        next_processing = next_target["processing"]

        if skip_to_step5:
            suffix = " (step 5부터)"
        elif skip_k:
            suffix = " (K 생략)"
        else:
            suffix = ""
        self.log(f"\n🔀 [{next_label}] 제작대로 전환 중...{suffix}")

        while True:  # 타임아웃 시 재시도 루프
            try:
                self.timeout_check_enabled = True

                if not skip_to_step5:
                    # --- 4. K 입력 → Processing1/2.png 클릭 ---
                    if not skip_k:
                        self.interruptible_sleep(1.5)
                        self.send_key_perfect(child_hwnd, self.VK_K, 0x25, 0.05)  # K key (scan code 0x25)
                        self.log("⌨️ 제작대 목록 열기")
                        self.interruptible_sleep(1.5)

                    # Processing1.png or Processing2.png 클릭 (유사도 0.7)
                    self.log("🔍 제작대 목록 탐색 중...")
                    found_processing = False
                    for proc_img in ['Processing1.png', 'Processing2.png']:
                        screen = self.get_window_screenshot(parent_hwnd)
                        pos = self.find_image_pos(screen, proc_img, threshold=0.7)
                        if pos:
                            self.background_click_pro(parent_hwnd, pos[0], pos[1])
                            self.log("✅ 제작대 목록 클릭")
                            found_processing = True
                            break
                    if not found_processing:
                        self.log("⚠️ 제작대 목록 미감지, 5초 대기 후 재시도...")
                        found_processing = self.find_and_click(parent_hwnd, 'Processing1.png', threshold=0.7, wait_time=5)
                        if not found_processing:
                            found_processing = self.find_and_click(parent_hwnd, 'Processing2.png', threshold=0.7, wait_time=5)
                    self.interruptible_sleep(1.5)

                # --- 5. 분야 Processing.png 클릭 ---
                self.log(f"🔍 {next_target['category']} 제작대 이동 중...")
                found = self.find_and_click(parent_hwnd, next_processing, threshold=0.7, wait_time=5)
                if found:
                    self.log(f"✅ {next_target['category']} 제작대 클릭")
                else:
                    self.log(f"⚠️ {next_target['category']} 제작대 미감지")
                self.interruptible_sleep(1.5)

                # --- 6. ReceiveAll 5초 감지 → 있으면 클릭 / 없으면 재료 클릭 → 스페이스바 ---
                self.log("🔍 모두 받기 버튼 5초 탐색 중...")
                receive_found = self.find_and_click(parent_hwnd, 'ReceiveAll.png', threshold=self.threshold, wait_time=5)

                if not receive_found:
                    # ReceiveAll 없음 → 재료 클릭 → 스페이스바
                    self.log(f"⏩ 모두 받기 없음 → {next_label} 클릭 시도")
                    screen = self.get_window_screenshot(parent_hwnd)
                    pos = self.find_image_pos(screen, next_img, threshold=self.threshold)
                    if pos:
                        self.background_click_pro(parent_hwnd, pos[0], pos[1])
                        self.log(f"⛏️ {next_label} 클릭")
                        self.interruptible_sleep(1.5)
                        self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                        self.log("⌨️ 제작 시작")
                        self.interruptible_sleep(1.5)

                # --- 7. Error 체크 ---
                error = self.check_error(parent_hwnd)
                if error == 'NoIngre':
                    # 7-1. 재료 부족 → 해당 재료 로테이션에서 제외
                    return "error"
                elif error == 'FullOrder':
                    # 7-1b. 주문 가득 참 → Esc 2회 후 다음 로테이션 step 5로
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
                    self.log("⌨️ 주문 가득 참 → 창 닫기 (1/2)")
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
                    self.log("⌨️ 주문 가득 참 → 창 닫기 (2/2)")
                    self.interruptible_sleep(1.5)
                    return "fullorder"

                if receive_found:
                    # 7-2. ReceiveAll이 클릭됐을 경우: 10초 대기 → ReceiveAll → 스페이스바
                    self.log("✅ 모두 받기 완료 → 10초 대기...")
                    self.interruptible_sleep(10.0)
                    self.log("🔍 모두 받기 버튼 재탐색 중...")
                    while self.is_running:
                        screen = self.get_window_screenshot(parent_hwnd)
                        pos = self.find_image_pos(screen, 'ReceiveAll.png', threshold=self.threshold)
                        if pos:
                            self.background_click_pro(parent_hwnd, pos[0], pos[1])
                            self.log("✅ 모두 받기 완료")
                            break
                        self.interruptible_sleep(1.0)
                    self.check_running()
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_SPACE, 0x39, 0.05)
                    self.log("⌨️ 확인")
                    self.interruptible_sleep(1.5)
                else:
                    # 7-3. 재료 클릭됐을 경우: 10초 대기 → Esc
                    self.log("⛏️ 제작 등록 완료 → 10초 대기...")
                    self.interruptible_sleep(10.0)
                    self.interruptible_sleep(1.5)
                    self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
                    self.log("⌨️ 제작대 닫기")
                    self.interruptible_sleep(1.5)

                self.log(f"✅ [{next_label}] 전환 완료")
                return "ok"

            except TimeoutRecoveryError as e:
                self.log(f"⏰ 타임아웃 감지: {e}")
                self.timeout_check_enabled = False
                if self.recovery_reset(parent_hwnd, child_hwnd):
                    self.log("🔄 Step 4(K키)부터 재시작합니다...")
                    skip_to_step5 = False
                    skip_k = False
                    continue  # Step 4(K키)부터 재시작
                else:
                    self.log("❌ 복구 실패 → 매크로 중지")
                    self.is_running = False
                    return "ok"
            finally:
                self.timeout_check_enabled = False

    def stop(self):
        self.is_running = False


class MainWindow(QWidget):
    """프로그램의 메인 화면(UI) — 모던 디자인"""
    def __init__(self):
        super().__init__()
        self.is_dark_mode = False
        self.colors = LIGHT_COLORS.copy()
        self.initUI()
        self.worker = None

    def initUI(self):
        self.setWindowTitle('모비노기 통합 제작 매크로')
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
        
        self.title_lbl = QLabel("모비노기 제작 매크로")
        self.title_lbl.setFont(QFont("Malgun Gothic", 13, QFont.Bold))
        
        header.addWidget(self.title_lbl)
        header.addStretch(1)

        self.btn_theme = QPushButton("🌙")
        self.btn_theme.setObjectName("btnTheme")
        self.btn_theme.setFixedSize(34, 34)
        self.btn_theme.setCursor(Qt.PointingHandCursor)
        self.btn_theme.setToolTip("다크모드 전환")
        self.btn_theme.clicked.connect(self.toggle_theme)
        header.addWidget(self.btn_theme)

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
        
        mode_title = QLabel("작업 모드 (작업대별 1개, 복수 선택 가능)")
        mode_title.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        mode_title.setObjectName("cardTitle")
        mode_card_layout.addWidget(mode_title)
        
        mode_grid = QGridLayout()
        mode_grid.setSpacing(8)
        self.checkboxes = []  # [(checkbox, cat_idx, item_idx), ...]

        for cat_idx, cat in enumerate(CATEGORIES):
            for item_idx, (label, img_file) in enumerate(cat["items"]):
                cb = QCheckBox(label)
                cb.img_file = img_file
                cb.cat_idx = cat_idx
                cb.processing = cat["processing"]
                cb.category_name = cat["name"]
                cb.setCursor(Qt.PointingHandCursor)
                cb.stateChanged.connect(lambda state, ci=cat_idx, ii=item_idx: self.on_checkbox_changed(ci, ii))
                self.checkboxes.append((cb, cat_idx, item_idx))
                mode_grid.addWidget(cb, cat_idx, item_idx)

        mode_card_layout.addLayout(mode_grid)
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
            "• 창 크기: 1280x960 설정 후 세로만 최소로 바꾸기\n"
            "• 반드시 선택한 재료의 제작대에 가서 제작대를 연 채로 시작\n"
            "• 이미지 스캔 방식이라 창이 가려지면 안 됩니다"
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
            
            /* 메모 입력 */
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
                selection-background-color: {c['accent']};
                selection-color: white;
                font-size: 10pt;
            }}
            
            /* 모드 카드 */
            QFrame#modeCard {{
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 12px;
            }}
            QLabel#cardTitle {{
                color: {c['accent']};
                border: none;
                background: transparent;
            }}
            
            /* 체크박스 */
            QCheckBox {{
                color: {c['text']};
                spacing: 6px;
                font-size: 10pt;
                background: transparent;
                border: none;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border-radius: 4px;
                border: 2px solid {c['text_dim']};
                background-color: transparent;
            }}
            QCheckBox::indicator:checked {{
                border: 2px solid {c['accent']};
                background-color: {c['accent']};
            }}
            QCheckBox::indicator:hover {{
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
        for w in [self.btn_start, self.btn_stop, self.btn_guide, self.btn_clear, self.log_text, self.btn_theme]:
            w.style().unpolish(w)
            w.style().polish(w)

    def on_checkbox_changed(self, cat_idx, item_idx):
        """같은 분야(행)에서 다른 항목이 체크되면 기존 체크 해제"""
        # 방금 변경된 체크박스 찾기
        changed_cb = None
        for cb, ci, ii in self.checkboxes:
            if ci == cat_idx and ii == item_idx:
                changed_cb = cb
                break
        
        if changed_cb and changed_cb.isChecked():
            # 같은 분야의 다른 체크박스 해제
            for cb, ci, ii in self.checkboxes:
                if ci == cat_idx and ii != item_idx:
                    cb.blockSignals(True)
                    cb.setChecked(False)
                    cb.blockSignals(False)

    def get_selected_targets(self):
        """선택된 체크박스들의 타겟 정보 리스트 반환"""
        targets = []
        for cb, cat_idx, item_idx in self.checkboxes:
            if cb.isChecked():
                targets.append({
                    "label": cb.text(),
                    "image": cb.img_file,
                    "category": cb.category_name,
                    "processing": cb.processing,
                })
        return targets

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        self.colors = DARK_COLORS.copy() if self.is_dark_mode else LIGHT_COLORS.copy()
        self.btn_theme.setText("☀️" if self.is_dark_mode else "🌙")
        self.btn_theme.setToolTip("라이트모드 전환" if self.is_dark_mode else "다크모드 전환")
        self.apply_style()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'btn_clear'):
            self.btn_clear.move(self.log_text.width() - 65, 5)

    def start_macro(self):
        targets = self.get_selected_targets()
        if not targets:
            self.append_log("⚠️ 최소 1개의 재료를 선택해주세요.")
            return

        self.worker = MacroWorker(targets)
        self.worker.log_signal.connect(self.append_log)
        self.worker.countdown_signal.connect(self.update_countdown)
        self.worker.finished_signal.connect(self.on_finished)
        
        self.thread = threading.Thread(target=self.worker.run)
        self.thread.daemon = True
        self.thread.start()
        
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        names = ", ".join([t["label"] for t in targets])
        mode = "로테이션" if len(targets) > 1 else "단일"
        self.append_log(f"=== {mode} 매크로 시작: {names} ===")

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
