import os
import random
import sys
import threading
import time
from dataclasses import dataclass
from functools import lru_cache
from typing import List, Optional, Tuple

import cv2
import ctypes
import numpy as np
import win32api
import win32con
import win32gui
from PIL import Image, ImageGrab
from PyQt5.QtCore import QObject, QSize, Qt, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QFont, QIcon, QImage, QPixmap, QTextBlockFormat, QTextCursor
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from pynput import keyboard as pynput_keyboard


kernel32 = ctypes.WinDLL("kernel32")
user32 = ctypes.WinDLL("user32")
hWnd = kernel32.GetConsoleWindow()
if hWnd:
    user32.ShowWindow(hWnd, win32con.SW_HIDE)

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

try:
    RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE_LANCZOS = Image.LANCZOS


WINDOW_TITLE = "마비노기 모바일"
TARGET_CLIENT_WIDTH = 1280
TARGET_CLIENT_HEIGHT = 691

FIRST_SLOT_X = 775
FIRST_SLOT_Y = 214
SLOT_DX = 100
SLOT_DY = 118
COLS = 5
ROWS = 3

NAME_OFFSET_X = -46
NAME_OFFSET_Y = 42
NAME_W = 95
NAME_H = 34

SLOT_MATCH_THRESHOLD = 0.88
BUTTON_THRESHOLD = 0.7
INGRE_THRESHOLD = 0.9
BUTTON_TIMEOUT = 10.0
BUTTON_RETRY_INTERVAL = 0.5
RECOVERY_MAX_ESC = 10
AUTO_WAIT_MIN = 90
AUTO_WAIT_MAX = 150

INVENTORY_OPEN_DELAY = 1.2
AFTER_SLOT_CLICK_DELAY = 0.4
AFTER_BUTTON_CLICK_DELAY = 0.5
AFTER_SELL_COMPLETE_DELAY = 1.5
AFTER_TAB_CLICK_DELAY = 0.6

VK_I = 0x49
SCAN_I = 0x17
VK_ESC = win32con.VK_ESCAPE
SCAN_ESC = 0x01

PROCESS_IMAGE = "Process.png"
INGRE_IMAGES = ["ingre1.png", "ingre2.png"]
BUTTON_SEQUENCE = [
    ("sell1.png", "판매 버튼1"),
    ("max.png", "최대 버튼"),
    ("sell2.png", "판매 버튼2"),
]
REQUIRED_IMAGES = [PROCESS_IMAGE] + [name for name, _ in BUTTON_SEQUENCE]


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


LIGHT_COLORS = {
    "bg": "#e7f2eb",
    "card": "#f8fcf9",
    "card_border": "#5f8f73",
    "surface": "#eef7f1",
    "text": "#1f3328",
    "text_dim": "#5c7768",
    "accent": "#2f9e66",
    "accent_hover": "#45b87c",
    "success": "#34b36f",
    "success_hover": "#4cc786",
    "danger": "#e07a7a",
    "danger_hover": "#ec9292",
    "utility": "#c98a2e",
    "utility_hover": "#d99b43",
    "log_bg": "#f4faf6",
    "log_text": "#2a5a43",
    "input_bg": "#f9fdfb",
    "input_border": "#5f8f73",
}

DARK_COLORS = {
    "bg": "#121b15",
    "card": "#1b2720",
    "card_border": "#3e6a52",
    "surface": "#1f2d25",
    "text": "#e4f2e8",
    "text_dim": "#9fbea9",
    "accent": "#53c083",
    "accent_hover": "#69cf95",
    "success": "#42b978",
    "success_hover": "#57c98a",
    "danger": "#d67979",
    "danger_hover": "#e08d8d",
    "utility": "#c79a49",
    "utility_hover": "#d7ab5a",
    "log_bg": "#142019",
    "log_text": "#8ce0b2",
    "input_bg": "#16231b",
    "input_border": "#3e6a52",
}


@dataclass
class SlotMatch:
    row: int
    col: int
    score: float


@dataclass
class ItemTemplate:
    label: str
    template_image: Image.Image
    source_row: int
    source_col: int
    created_at: float


@dataclass
class SellCandidate:
    template: ItemTemplate
    match: SlotMatch


class TimeoutRecoveryError(Exception):
    pass


def resource_path(name: str) -> str:
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, name)


def slot_to_text(row: int, col: int) -> str:
    return f"{row + 1}행 {col + 1}열"


def clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def normalize_capture_key(value: str) -> Optional[str]:
    text = (value or "").strip()
    if not text:
        return None
    key = text[0]
    if not key.isprintable() or key.isspace():
        return None
    return key.lower()


def format_capture_key(value: str) -> str:
    return (value or "q").upper()


def image_similarity(img1: Image.Image, img2: Image.Image) -> float:
    a = pil_to_bgr(img1)
    b = pil_to_bgr(img2)
    result = cv2.matchTemplate(a, b, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, _ = cv2.minMaxLoc(result)
    return float(max_val)


def get_slot_icon_center(row: int, col: int) -> Tuple[int, int]:
    rel_x = FIRST_SLOT_X + col * SLOT_DX
    rel_y = FIRST_SLOT_Y + row * SLOT_DY
    return rel_x, rel_y


def get_name_rect_relative(row: int, col: int) -> Tuple[int, int, int, int]:
    cx, cy = get_slot_icon_center(row, col)
    return cx + NAME_OFFSET_X, cy + NAME_OFFSET_Y, NAME_W, NAME_H


def crop_relative_from_client(client_img: Image.Image, x: int, y: int, w: int, h: int) -> Image.Image:
    return client_img.crop((x, y, x + w, y + h))


def wait_for_capture_key(trigger_key: str) -> Tuple[str, Optional[Tuple[int, int]]]:
    capture_key = normalize_capture_key(trigger_key) or "q"
    result = {"action": "cancel", "cursor_pos": None}

    def on_press(key):
        try:
            if key.char and key.char.lower() == capture_key:
                result["action"] = "capture"
                result["cursor_pos"] = win32api.GetCursorPos()
                return False
        except AttributeError:
            if key == pynput_keyboard.Key.esc:
                result["action"] = "cancel"
                return False
        return True

    with pynput_keyboard.Listener(on_press=on_press) as listener:
        listener.join()
    return result["action"], result["cursor_pos"]


def pil_to_bgr(img: Image.Image) -> np.ndarray:
    rgb = np.array(img.convert("RGB"))
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def pil_image_to_icon(img: Image.Image, max_width: int = 72, max_height: int = 26) -> QIcon:
    rgba = img.convert("RGBA")
    scale = min(max_width / rgba.width, max_height / rgba.height)
    new_width = max(1, int(rgba.width * scale))
    new_height = max(1, int(rgba.height * scale))
    resized = rgba.resize((new_width, new_height), RESAMPLE_LANCZOS)

    canvas = Image.new("RGBA", (max_width, max_height), (0, 0, 0, 0))
    offset_x = (max_width - new_width) // 2
    offset_y = (max_height - new_height) // 2
    canvas.paste(resized, (offset_x, offset_y))

    data = canvas.tobytes("raw", "RGBA")
    qimage = QImage(data, canvas.width, canvas.height, QImage.Format_RGBA8888).copy()
    pixmap = QPixmap.fromImage(qimage)
    return QIcon(pixmap)


def resize_client_area_for_hwnd(hwnd: int) -> Tuple[int, int]:
    if not hwnd:
        raise RuntimeError("창을 찾지 못했습니다.")

    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.3)

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    has_menu = bool(win32gui.GetMenu(hwnd))

    rect = RECT(0, 0, TARGET_CLIENT_WIDTH, TARGET_CLIENT_HEIGHT)
    ctypes.windll.user32.AdjustWindowRectEx(
        ctypes.byref(rect),
        style,
        has_menu,
        ex_style,
    )

    outer_width = rect.right - rect.left
    outer_height = rect.bottom - rect.top
    win32gui.MoveWindow(hwnd, left, top, outer_width, outer_height, True)

    client_left, client_top, client_right, client_bottom = win32gui.GetClientRect(hwnd)
    client_width = client_right - client_left
    client_height = client_bottom - client_top
    return client_width, client_height


@lru_cache(maxsize=None)
def load_cv_template(image_name: str) -> Optional[np.ndarray]:
    full_path = resource_path(image_name)
    if not os.path.exists(full_path):
        return None
    try:
        img_array = np.fromfile(full_path, np.uint8)
        template = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    except Exception:
        return None
    return template


class NoticeDialog(QDialog):
    def __init__(self, colors, parent=None):
        super().__init__(parent)
        self.colors = colors
        self.init_ui(parent)

    def init_ui(self, parent):
        self.setWindowTitle("모비노기 채집물 자동 판매 가이드")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setFixedSize(400, 560)

        icon_path = resource_path("my_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        if parent:
            self.setGeometry(parent.geometry())

        c = self.colors
        self.setStyleSheet(
            f"""
            QDialog {{
                background-color: {c['bg']};
            }}
            """
        )

        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        notice_text = QTextBrowser()
        notice_text.setStyleSheet(
            f"""
            QTextBrowser {{
                background-color: {c['card']};
                color: {c['text']};
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                border: 1px solid {c['card_border']};
                border-radius: 10px;
                padding: 12px;
            }}
            """
        )
        notice_text.setHtml(
            """
            <h3 style='color: #6c5ce7;'>[ 필독: 매크로 가이드 ]</h3>
            <p>이 매크로는 인벤토리를 스캔해서, 미리 지정한 재료 아이템을 자동으로 찾아 판매합니다.</p>
            <hr>
            <p><b>1. 판매 아이템 선택하기</b></p>
            <ul>
                <li><b>판매 아이템 선택</b> 버튼을 눌러요.</li>
                <li>인벤토리를 열고 판매 할 아이템 위에 마우스를 올린 뒤 <b>키보드 Q</b>를 누르기.</li>
                <li>여러 종류를 팔고 싶으면 여러번 반복하면 돼요.</li>
                <li>재료 아이템만 가능하니까 유의해주세요</li>
            </ul>
            <p><b>2. 자동 판매</b></p>
            <ul>
                <li><b>자동 판매 시작</b>을 누르면 등록한 아이템들을 인벤토리에서 스캔해요.</li>
                <li>인벤토리에서 판매할 아이템을 찾으면 자동으로 판매해요.</li>
                <li>에러 탈출 로직이 포함되어 있고, 랜덤한 시간마다 반복해서 판매를 진행해요.</li>
            </ul>
            <p><b>3. 주의사항</b></p>
            <ul>
                <li>프로그램 시작 시 게임 창 크기를 한 번 자동으로 맞춥니다.</li>
                <li>이미지 스캔 방식이라 마비노기 창이 가려지면 안돼요.</li>
                <li>창 크기가 변경되었을 경우 정상적으로 작동하지 않을 수 있어요..</li>
                <li>창 크기를 실수로 변경했다면 <b>창 크기 맞추기</b> 버튼을 눌러주세요.</li>
            </ul>
            """
        )

        btn_close = QPushButton("확인했습니다")
        btn_close.setFixedHeight(44)
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet(
            f"""
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
            """
        )
        btn_close.clicked.connect(self.accept)

        layout.addWidget(notice_text)
        layout.addWidget(btn_close)
        self.setLayout(layout)


class NoticeDialog(QDialog):
    def __init__(self, colors, capture_key: str, parent=None):
        super().__init__(parent)
        self.colors = colors
        self.capture_key = capture_key
        self.init_ui(parent)

    def init_ui(self, parent):
        self.setWindowTitle("모비노기 채집물 자동 판매 가이드")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setFixedSize(400, 560)

        icon_path = resource_path("my_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        if parent:
            self.setGeometry(parent.geometry())

        c = self.colors
        self.setStyleSheet(
            f"""
            QDialog {{
                background-color: {c['bg']};
            }}
            """
        )

        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        notice_text = QTextBrowser()
        notice_text.setStyleSheet(
            f"""
            QTextBrowser {{
                background-color: {c['card']};
                color: {c['text']};
                font-family: 'Malgun Gothic';
                font-size: 10pt;
                border: 1px solid {c['card_border']};
                border-radius: 10px;
                padding: 12px;
            }}
            """
        )
        display_key = format_capture_key(self.capture_key)
        notice_text.setHtml(
            f"""
            <h3 style='color: #2f9e66;'>[ 안내: 자동 판매 매크로 가이드 ]</h3>
            <p>이 매크로는 인벤토리를 검사해서 미리 등록한 판매 아이템을 자동으로 찾아 판매합니다.</p>
            <hr>
            <p><b>1. 판매 아이템 선택</b></p>
            <ul>
                <li><b>판매 아이템 선택</b> 버튼을 누릅니다.</li>
                <li>인벤토리를 열고 판매 할 아이템 위에 마우스를 올린 뒤 <b>{display_key}</b> 키를 누릅니다.</li>
                <li>여러 품목을 등록하려면 여러번 반복하면 됩니다.</li>
                <li>선택 모드 취소는 <b>Esc</b>입니다.</li>
            </ul>
            <p><b>2. 자동 판매</b></p>
            <ul>
                <li><b>자동 판매 시작</b>을 누르면 자동으로 인벤토리를 스캔하여 해당 아이템을 판매합니다..</li>
                <li>에러 탈출 로직이 포함되어 있으며, 랜덤한 시간마다 반복해서 판매를 진행합니다.</li>
            </ul>
            <p><b>3. 주의 사항</b></p>
            <ul>
                <li>프로그램 시작 시 게임 창 크기를 자동으로 한 번 맞춥니다.</li>
                <li>이미지 스캔 방식이라 게임 창 크기가 바뀌면 정확도가 떨어질 수 있습니다.</li>
                <li>창 크기를 바꿨다면 <b>창 크기 맞추기</b>를 다시 눌러 주세요.</li>
                <li>설정 버튼에서 선택 키를 바꿀 수 있습니다.</li>
            </ul>
            """
        )

        btn_close = QPushButton("확인했습니다")
        btn_close.setFixedHeight(44)
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet(
            f"""
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
            """
        )
        btn_close.clicked.connect(self.accept)

        layout.addWidget(notice_text)
        layout.addWidget(btn_close)
        self.setLayout(layout)


class SettingsPopup(QDialog):
    def __init__(self, owner):
        super().__init__(owner, Qt.Popup | Qt.FramelessWindowHint)
        self.owner = owner
        self.setFixedSize(248, 144)
        self.init_ui()
        self.refresh_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        self.title_lbl = QLabel("설정")
        self.title_lbl.setFont(QFont("Malgun Gothic", 10, QFont.Bold))
        self.title_lbl.setObjectName("settingsTitle")


        key_label = QLabel("판매 아이템 선택 키")
        key_label.setObjectName("settingsLabel")
        key_label.setFont(QFont("Malgun Gothic", 8))

        key_row = QHBoxLayout()
        key_row.setSpacing(8)

        self.key_input = QLineEdit()
        self.key_input.setObjectName("settingsInput")
        self.key_input.setMaxLength(1)
        self.key_input.setFixedSize(52, 34)
        self.key_input.setAlignment(Qt.AlignCenter)
        self.key_input.returnPressed.connect(self.apply_capture_key)

        self.btn_apply_key = QPushButton("저장")
        self.btn_apply_key.setObjectName("settingsApply")
        self.btn_apply_key.setCursor(Qt.PointingHandCursor)
        self.btn_apply_key.setFixedHeight(34)
        self.btn_apply_key.clicked.connect(self.apply_capture_key)

        key_row.addWidget(self.key_input)
        key_row.addWidget(self.btn_apply_key, 1)

        self.hint_lbl = QLabel("한 글자 키만 설정할 수 있습니다.")
        self.hint_lbl.setObjectName("settingsHint")
        self.hint_lbl.setWordWrap(True)
        self.hint_lbl.setFont(QFont("Malgun Gothic", 8))

        layout.addWidget(self.title_lbl)
        layout.addWidget(key_label)
        layout.addLayout(key_row)
        layout.addWidget(self.hint_lbl)

    def refresh_ui(self):
        c = self.owner.colors
        self.setStyleSheet(
            f"""
            QDialog {{
                background-color: {c['surface']};
                border: 2px solid {c['accent']};
                border-radius: 14px;
            }}
            QLabel {{
                color: {c['text']};
                background: transparent;
                border: none;
            }}
            QLabel#settingsTitle {{
                color: {c['accent']};
                font-size: 11pt;
                font-weight: bold;
                padding-bottom: 2px;
            }}
            QLabel#settingsLabel {{
                color: {c['text']};
                font-weight: bold;
            }}
            QLabel#settingsHint {{
                color: {c['text_dim']};
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                padding: 6px 8px;
            }}
            QPushButton {{
                background-color: {c['card']};
                color: {c['text']};
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                border: 1px solid {c['accent']};
            }}
            QPushButton#settingsApply {{
                background-color: {c['accent']};
                color: white;
                border: 1px solid {c['accent']};
            }}
            QPushButton#settingsApply:hover {{
                background-color: {c['accent_hover']};
                border: 1px solid {c['accent_hover']};
            }}
            QLineEdit#settingsInput {{
                background-color: {c['input_bg']};
                color: {c['text']};
                border: 2px solid {c['accent']};
                border-radius: 8px;
                padding: 2px 8px;
                font-size: 12pt;
                font-weight: bold;
            }}
            """
        )
        self.btn_theme_toggle.setText("다크모드: ON" if self.owner.is_dark_mode else "다크모드: OFF")
        self.key_input.setText(format_capture_key(self.owner.capture_key))
        self.key_input.selectAll()

    def on_toggle_theme(self):
        self.owner.toggle_theme()
        self.refresh_ui()

    def apply_capture_key(self):
        value = normalize_capture_key(self.key_input.text())
        if value is None:
            self.key_input.setText(format_capture_key(self.owner.capture_key))
            self.key_input.selectAll()
            self.owner.append_log("설정 실패: 판매 아이템 선택 키는 한 글자만 사용할 수 있습니다.")
            return
        self.owner.set_capture_key(value)
        self.refresh_ui()


class SettingsPopup(QDialog):
    def __init__(self, owner):
        super().__init__(owner, Qt.Popup | Qt.FramelessWindowHint)
        self.owner = owner
        self.setFixedSize(220, 118)
        self.init_ui()
        self.refresh_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        self.title_lbl = QLabel("설정")
        self.title_lbl.setFont(QFont("Malgun Gothic", 10, QFont.Bold))

        key_label = QLabel("판매 아이템 선택 키")
        key_label.setObjectName("settingsLabel")
        key_label.setFont(QFont("Malgun Gothic", 8))

        key_row = QHBoxLayout()
        key_row.setSpacing(8)

        self.key_input = QLineEdit()
        self.key_input.setMaxLength(1)
        self.key_input.setFixedSize(42, 30)
        self.key_input.setAlignment(Qt.AlignCenter)
        self.key_input.returnPressed.connect(self.apply_capture_key)

        self.btn_apply_key = QPushButton("저장")
        self.btn_apply_key.setCursor(Qt.PointingHandCursor)
        self.btn_apply_key.setFixedHeight(30)
        self.btn_apply_key.clicked.connect(self.apply_capture_key)

        key_row.addWidget(self.key_input)
        key_row.addWidget(self.btn_apply_key, 1)

        self.hint_lbl = QLabel("한 글자 키만 설정할 수 있습니다.")
        self.hint_lbl.setObjectName("settingsHint")
        self.hint_lbl.setWordWrap(True)
        self.hint_lbl.setFont(QFont("Malgun Gothic", 8))

        layout.addWidget(self.title_lbl)
        layout.addWidget(key_label)
        layout.addLayout(key_row)
        layout.addWidget(self.hint_lbl)

    def refresh_ui(self):
        c = self.owner.colors
        self.setStyleSheet(
            f"""
            QDialog {{
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 12px;
            }}
            QLabel {{
                color: {c['text']};
                background: transparent;
                border: none;
            }}
            QLabel#settingsHint {{
                color: {c['text_dim']};
            }}
            QPushButton {{
                background-color: {c['surface']};
                color: {c['text']};
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                border: 1px solid {c['accent']};
            }}
            QLineEdit {{
                background-color: {c['input_bg']};
                color: {c['text']};
                border: 1px solid {c['input_border']};
                border-radius: 8px;
                padding: 2px 6px;
                font-size: 11pt;
                font-weight: bold;
            }}
            """
        )
        self.key_input.setText(format_capture_key(self.owner.capture_key))
        self.key_input.selectAll()

    def apply_capture_key(self):
        value = normalize_capture_key(self.key_input.text())
        if value is None:
            self.key_input.setText(format_capture_key(self.owner.capture_key))
            self.key_input.selectAll()
            self.owner.append_log("설정 실패: 판매 아이템 선택 키는 한 글자만 사용할 수 있습니다.")
            return
        self.owner.set_capture_key(value)
        self.refresh_ui()


class SettingsPopup(QDialog):
    def __init__(self, owner):
        super().__init__(owner, Qt.Popup | Qt.FramelessWindowHint)
        self.owner = owner
        self.setFixedSize(248, 144)
        self.init_ui()
        self.refresh_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        self.title_lbl = QLabel("설정")
        self.title_lbl.setObjectName("settingsTitle")
        self.title_lbl.setFont(QFont("Malgun Gothic", 10, QFont.Bold))

        key_label = QLabel("판매 아이템 선택 키")
        key_label.setObjectName("settingsLabel")
        key_label.setFont(QFont("Malgun Gothic", 8))

        key_row = QHBoxLayout()
        key_row.setSpacing(8)

        self.key_input = QLineEdit()
        self.key_input.setObjectName("settingsInput")
        self.key_input.setMaxLength(1)
        self.key_input.setFixedSize(52, 34)
        self.key_input.setAlignment(Qt.AlignCenter)
        self.key_input.returnPressed.connect(self.apply_capture_key)

        self.btn_apply_key = QPushButton("저장")
        self.btn_apply_key.setObjectName("settingsApply")
        self.btn_apply_key.setCursor(Qt.PointingHandCursor)
        self.btn_apply_key.setFixedHeight(34)
        self.btn_apply_key.clicked.connect(self.apply_capture_key)

        key_row.addWidget(self.key_input)
        key_row.addWidget(self.btn_apply_key, 1)

        self.hint_lbl = QLabel("한 글자 키만 설정할 수 있습니다.")
        self.hint_lbl.setObjectName("settingsHint")
        self.hint_lbl.setWordWrap(True)
        self.hint_lbl.setFont(QFont("Malgun Gothic", 8))

        layout.addWidget(self.title_lbl)
        layout.addWidget(key_label)
        layout.addLayout(key_row)
        layout.addWidget(self.hint_lbl)

    def refresh_ui(self):
        c = self.owner.colors
        self.setStyleSheet(
            f"""
            QDialog {{
                background-color: {c['surface']};
                border: 2px solid {c['accent']};
                border-radius: 14px;
            }}
            QLabel {{
                color: {c['text']};
                background: transparent;
                border: none;
            }}
            QLabel#settingsTitle {{
                color: {c['accent']};
                font-size: 11pt;
                font-weight: bold;
                padding-bottom: 2px;
            }}
            QLabel#settingsLabel {{
                color: {c['text']};
                font-weight: bold;
            }}
            QLabel#settingsHint {{
                color: {c['text_dim']};
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                padding: 6px 8px;
            }}
            QPushButton {{
                background-color: {c['card']};
                color: {c['text']};
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                border: 1px solid {c['accent']};
            }}
            QPushButton#settingsApply {{
                background-color: {c['accent']};
                color: white;
                border: 1px solid {c['accent']};
            }}
            QPushButton#settingsApply:hover {{
                background-color: {c['accent_hover']};
                border: 1px solid {c['accent_hover']};
            }}
            QLineEdit#settingsInput {{
                background-color: {c['input_bg']};
                color: {c['text']};
                border: 2px solid {c['accent']};
                border-radius: 8px;
                padding: 2px 8px;
                font-size: 12pt;
                font-weight: bold;
            }}
            """
        )
        self.key_input.setText(format_capture_key(self.owner.capture_key))
        self.key_input.selectAll()

    def apply_capture_key(self):
        value = normalize_capture_key(self.key_input.text())
        if value is None:
            self.key_input.setText(format_capture_key(self.owner.capture_key))
            self.key_input.selectAll()
            self.owner.append_log("설정 실패: 판매 아이템 선택 키는 한 글자만 사용할 수 있습니다.")
            return
        self.owner.set_capture_key(value)
        self.refresh_ui()


class BaseWorker(QObject):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    countdown_signal = pyqtSignal(int)
    template_selected_signal = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self.is_running = False

    def log(self, message: str) -> None:
        self.log_signal.emit(message)

    def stop(self) -> None:
        self.is_running = False

    def check_running(self) -> None:
        if not self.is_running:
            raise InterruptedError("중지됨")

    def interruptible_sleep(self, seconds: float) -> None:
        end_time = time.time() + max(0.0, seconds)
        while time.time() < end_time:
            self.check_running()
            time.sleep(min(0.1, end_time - time.time()))

    def get_window_handles(self) -> Tuple[int, int]:
        parent_hwnd = win32gui.FindWindow(None, WINDOW_TITLE)
        if not parent_hwnd:
            raise RuntimeError(f"'{WINDOW_TITLE}' 창을 찾을 수 없습니다.")
        input_hwnd = win32gui.FindWindowEx(parent_hwnd, None, None, None) or parent_hwnd
        return parent_hwnd, input_hwnd

    def restore_window(self, hwnd: int) -> None:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)

    def get_client_origin(self, hwnd: int) -> Tuple[int, int]:
        return win32gui.ClientToScreen(hwnd, (0, 0))

    def get_client_size(self, hwnd: int) -> Tuple[int, int]:
        left, top, right, bottom = win32gui.GetClientRect(hwnd)
        return right - left, bottom - top

    def capture_client_image(self, hwnd: int) -> Image.Image:
        client_left, client_top = self.get_client_origin(hwnd)
        client_width, client_height = self.get_client_size(hwnd)
        bbox = (
            client_left,
            client_top,
            client_left + client_width,
            client_top + client_height,
        )
        return ImageGrab.grab(bbox=bbox)

    def get_nearest_slot_from_mouse(
        self,
        hwnd: int,
        cursor_pos: Optional[Tuple[int, int]] = None,
    ) -> Tuple[int, int]:
        client_left, client_top = self.get_client_origin(hwnd)
        if cursor_pos is None:
            mouse_x, mouse_y = win32api.GetCursorPos()
        else:
            mouse_x, mouse_y = cursor_pos
        rel_x = mouse_x - client_left
        rel_y = mouse_y - client_top
        col = round((rel_x - FIRST_SLOT_X) / SLOT_DX)
        row = round((rel_y - FIRST_SLOT_Y) / SLOT_DY)
        return clamp(row, 0, ROWS - 1), clamp(col, 0, COLS - 1)

    def build_slot_template(self, client_img: Image.Image, row: int, col: int) -> Image.Image:
        x, y, w, h = get_name_rect_relative(row, col)
        return crop_relative_from_client(client_img, x, y, w, h)

    def scan_slot_matches(
        self,
        client_img: Image.Image,
        template_img: Image.Image,
    ) -> Tuple[List[SlotMatch], List[SlotMatch]]:
        all_matches: List[SlotMatch] = []
        valid_matches: List[SlotMatch] = []

        for row in range(ROWS):
            for col in range(COLS):
                slot_img = self.build_slot_template(client_img, row, col)
                score = image_similarity(template_img, slot_img)
                match = SlotMatch(row=row, col=col, score=score)
                all_matches.append(match)
                if score >= SLOT_MATCH_THRESHOLD:
                    valid_matches.append(match)

        all_matches.sort(key=lambda item: item.score, reverse=True)
        valid_matches.sort(key=lambda item: item.score, reverse=True)
        return all_matches, valid_matches

    def send_key_perfect(self, hwnd: int, vk_code: int, scan_code: int, duration: float = 0.05) -> None:
        self.check_running()
        win32gui.PostMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        down_lparam = 1 | (scan_code << 16)
        up_lparam = 1 | (scan_code << 16) | (1 << 30) | (1 << 31)
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, down_lparam)
        self.interruptible_sleep(duration)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, up_lparam)

    def click_client_point(self, hwnd: int, x: int, y: int) -> None:
        self.check_running()
        old_pos = win32api.GetCursorPos()
        client_left, client_top = self.get_client_origin(hwnd)
        screen_x = client_left + int(x)
        screen_y = client_top + int(y)
        try:
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_CLICKACTIVE, 0)
            win32api.SetCursorPos((screen_x, screen_y))
            self.interruptible_sleep(0.05)
            ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            self.interruptible_sleep(0.08)
            ctypes.windll.user32.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        finally:
            self.interruptible_sleep(0.05)
            win32api.SetCursorPos(old_pos)

    def find_image_pos(
        self,
        client_img: Image.Image,
        image_name: str,
        threshold: float = BUTTON_THRESHOLD,
    ) -> Optional[Tuple[int, int, float]]:
        template = load_cv_template(image_name)
        if template is None:
            return None

        screen = pil_to_bgr(client_img)
        try:
            result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
        except Exception:
            return None

        if max_val < threshold:
            return None

        th, tw = template.shape[:2]
        return max_loc[0] + tw // 2, max_loc[1] + th // 2, float(max_val)

    def find_and_click(
        self,
        hwnd: int,
        image_name: str,
        threshold: float = BUTTON_THRESHOLD,
        wait_time: float = BUTTON_TIMEOUT,
        label: Optional[str] = None,
    ) -> None:
        display = label or image_name
        start_time = time.time()
        while time.time() - start_time < wait_time:
            self.check_running()
            client_img = self.capture_client_image(hwnd)
            pos = self.find_image_pos(client_img, image_name, threshold)
            if pos:
                self.click_client_point(hwnd, pos[0], pos[1])
                self.log(f"✅ {display} 클릭")
                return
            self.interruptible_sleep(BUTTON_RETRY_INTERVAL)
        raise TimeoutRecoveryError(f"{display} {wait_time:.0f}초간 미감지")

    def find_and_click_any(
        self,
        hwnd: int,
        image_names: List[str],
        threshold: float = BUTTON_THRESHOLD,
        wait_time: float = BUTTON_TIMEOUT,
        label: Optional[str] = None,
    ) -> None:
        display = label or " / ".join(image_names)
        start_time = time.time()
        while time.time() - start_time < wait_time:
            self.check_running()
            client_img = self.capture_client_image(hwnd)
            for image_name in image_names:
                pos = self.find_image_pos(client_img, image_name, threshold)
                if pos:
                    self.click_client_point(hwnd, pos[0], pos[1])
                    self.log(f"✅ {display} 클릭")
                    return
            self.interruptible_sleep(BUTTON_RETRY_INTERVAL)
        raise TimeoutRecoveryError(f"{display} {wait_time:.0f}초간 미감지")

    def is_process_visible(self, hwnd: int) -> bool:
        client_img = self.capture_client_image(hwnd)
        return self.find_image_pos(client_img, PROCESS_IMAGE, BUTTON_THRESHOLD) is not None

    def recover_to_process(self, parent_hwnd: int, input_hwnd: int) -> bool:
        self.log("═══ 복구 시작 ═══")
        for attempt in range(1, RECOVERY_MAX_ESC + 1):
            self.check_running()
            self.send_key_perfect(input_hwnd, VK_ESC, SCAN_ESC, 0.05)
            self.log(f"복구 시도")
            self.interruptible_sleep(1.2)
            if self.is_process_visible(parent_hwnd):
                self.log("✅ 메인화면 감지 — 복구 완료")
                self.log("═══ 복구 종료 ═══")
                return True
        self.log("❌ 메인화면 미감지 — 복구 실패")
        self.log("═══ 복구 종료 ═══")
        return False


class ItemSelectionWorker(BaseWorker):
    def __init__(self, label_index: int, capture_key: str):
        super().__init__()
        self.label_index = label_index
        self.capture_key = capture_key

    def run(self) -> None:
        self.is_running = True
        try:
            parent_hwnd, _ = self.get_window_handles()
            self.restore_window(parent_hwnd)

            client_left, client_top = self.get_client_origin(parent_hwnd)
            client_width, client_height = self.get_client_size(parent_hwnd)

            action, cursor_pos = wait_for_capture_key(self.capture_key)
            if action != "capture" or cursor_pos is None:
                self.log("선택이 취소되었습니다.")
                return

            client_img = self.capture_client_image(parent_hwnd)

            selected_row, selected_col = self.get_nearest_slot_from_mouse(parent_hwnd, cursor_pos)
            template_img = self.build_slot_template(client_img, selected_row, selected_col)
            all_matches, valid_matches = self.scan_slot_matches(client_img, template_img)

            if not valid_matches:
                self.log("찾지 못했습니다. 그래도 선택한 아이템은 캐시에 등록합니다.")
            else:
                for match in valid_matches:
                    pass

            best_match = valid_matches[0] if valid_matches else all_matches[0]

            template = ItemTemplate(
                label=f"판매 아이템 {self.label_index}",
                template_image=template_img.copy(),
                source_row=selected_row,
                source_col=selected_col,
                created_at=time.time(),
            )
            self.template_selected_signal.emit(template)
            self.log(f"✅ {template.label} 등록 완료")
        except InterruptedError:
            self.log("선택 작업이 중지되었습니다.")
        except Exception as exc:
            self.log(f"❌ 선택 중 오류: {exc}")
        finally:
            self.is_running = False
            self.finished_signal.emit()


class AutoSellWorker(BaseWorker):
    def __init__(self, templates: List[ItemTemplate]):
        super().__init__()
        self.templates = templates

    def validate_assets(self) -> None:
        missing = [name for name in REQUIRED_IMAGES if load_cv_template(name) is None]
        if missing:
            raise RuntimeError("필수 이미지가 없습니다: " + ", ".join(missing))
        if not any(load_cv_template(name) is not None for name in INGRE_IMAGES):
            raise RuntimeError("재료 탭 이미지가 없습니다: ingre1.png 또는 ingre2.png")

    def ensure_base_screen(self, parent_hwnd: int, input_hwnd: int) -> None:
        if self.is_process_visible(parent_hwnd):
            return
        self.log("기준 화면이 아니어서 복구를 시도합니다.")
        if not self.recover_to_process(parent_hwnd, input_hwnd):
            raise TimeoutRecoveryError("기준 화면 복구 실패")

    def find_best_slot_for_template_in_image(
        self,
        client_img: Image.Image,
        template: ItemTemplate,
    ) -> Optional[SlotMatch]:
        _, valid_matches = self.scan_slot_matches(client_img, template.template_image)
        if not valid_matches:
            return None
        return valid_matches[0]

    def collect_current_candidates(self, parent_hwnd: int) -> List[SellCandidate]:
        client_img = self.capture_client_image(parent_hwnd)
        slot_best_candidates = {}

        for template in self.templates:
            _, valid_matches = self.scan_slot_matches(client_img, template.template_image)
            for match in valid_matches:
                key = (match.row, match.col)
                candidate = SellCandidate(template=template, match=match)
                existing = slot_best_candidates.get(key)
                if existing is None or candidate.match.score > existing.match.score:
                    slot_best_candidates[key] = candidate

        return list(slot_best_candidates.values())

    def pick_next_candidate(self, candidates: List[SellCandidate]) -> SellCandidate:
        return max(
            candidates,
            key=lambda candidate: (
                candidate.match.row,
                candidate.match.col,
                candidate.match.score,
            ),
        )

    def focus_ingredient_tab(self, parent_hwnd: int) -> None:
        self.find_and_click_any(parent_hwnd, INGRE_IMAGES, INGRE_THRESHOLD, BUTTON_TIMEOUT, "재료 탭")
        self.interruptible_sleep(AFTER_TAB_CLICK_DELAY)

    def run_cycle(self, parent_hwnd: int, input_hwnd: int) -> None:
        self.log("\n🔄 자동 판매 사이클 시작...")
        self.ensure_base_screen(parent_hwnd, input_hwnd)

        self.send_key_perfect(input_hwnd, VK_I, SCAN_I, 0.05)
        self.log("인벤토리 열기")
        self.interruptible_sleep(INVENTORY_OPEN_DELAY)
        self.focus_ingredient_tab(parent_hwnd)

        sold_count = 0
        try:
            while True:
                self.check_running()
                self.log("🔍 현재 인벤토리 전체 재검사")
                candidates = self.collect_current_candidates(parent_hwnd)
                if not candidates:
                    self.log("➖ 현재 화면에서 판매 대상이 더 이상 없습니다")
                    break

                candidate = self.pick_next_candidate(candidates)
                template = candidate.template
                match = candidate.match

                self.log(
                    f"🎯 이번 판매 대상: {template.label}"
                )
                slot_x, slot_y = get_slot_icon_center(match.row, match.col)
                self.click_client_point(parent_hwnd, slot_x, slot_y)
                self.interruptible_sleep(AFTER_SLOT_CLICK_DELAY)

                for image_name, label in BUTTON_SEQUENCE:
                    self.find_and_click(parent_hwnd, image_name, BUTTON_THRESHOLD, BUTTON_TIMEOUT, label)
                    self.interruptible_sleep(AFTER_BUTTON_CLICK_DELAY)

                sold_count += 1
                self.log(f"✅ {template.label}: 판매 완료")
                self.interruptible_sleep(AFTER_SELL_COMPLETE_DELAY)
                self.log("↻ 판매 후 인벤토리 재검사 준비")
        finally:
            self.send_key_perfect(input_hwnd, VK_ESC, SCAN_ESC, 0.05)
            self.log("인벤토리 닫기")
            self.interruptible_sleep(0.8)

        wait_next = random.randint(AUTO_WAIT_MIN, AUTO_WAIT_MAX)
        self.log(f"🏁 사이클 완료. 처리 성공 {sold_count}건 / {wait_next}초 후 재시작")
        for seconds in range(wait_next, -1, -1):
            self.check_running()
            self.countdown_signal.emit(seconds)
            self.interruptible_sleep(1.0)
        self.countdown_signal.emit(-1)

    def run(self) -> None:
        self.is_running = True
        try:
            if not self.templates:
                self.log("❌ 판매 아이템이 선택되지 않았습니다.")
                return

            self.validate_assets()
            parent_hwnd, input_hwnd = self.get_window_handles()
            self.restore_window(parent_hwnd)

            self.log(f"=== 자동 판매 시작 / 등록된 아이템 {len(self.templates)}개 ===")

            while self.is_running:
                try:
                    self.run_cycle(parent_hwnd, input_hwnd)
                except TimeoutRecoveryError as exc:
                    self.log(f"⏰ 타임아웃: {exc}")
                    if self.recover_to_process(parent_hwnd, input_hwnd):
                        self.log("🔄 복구 완료 — 다음 사이클로 진행합니다.")
                    else:
                        self.log("❌ 복구 실패 — 자동 판매를 중지합니다.")
                        break
                except InterruptedError:
                    raise
                except Exception as exc:
                    self.log(f"⚠️ 실행 중 오류: {exc}")
                    break
        except InterruptedError:
            self.log("🛑 자동 판매가 중지되었습니다.")
        except Exception as exc:
            self.log(f"❌ 시작 실패: {exc}")
        finally:
            self.is_running = False
            self.countdown_signal.emit(-1)
            self.finished_signal.emit()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.base_window_height = 560
        self.expanded_window_height = 640
        self.is_dark_mode = True
        self.colors = DARK_COLORS.copy()
        self.capture_key = "q"
        self.selected_templates: List[ItemTemplate] = []
        self.selection_worker: Optional[ItemSelectionWorker] = None
        self.auto_worker: Optional[AutoSellWorker] = None
        self.selection_thread = None
        self.auto_thread = None
        self.is_selecting = False
        self.is_auto_running = False
        self.is_settings_open = False
        self.settings_popup: Optional[SettingsPopup] = None
        self.init_ui()
        self.resize_game_window(auto=True)

    def init_ui(self):
        self.setWindowTitle("모비노기 채집물 자동 판매")
        self.setFixedSize(400, self.base_window_height)

        icon_path = resource_path("my_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(16, 14, 16, 12)
        self.main_layout.setSpacing(10)

        header = QHBoxLayout()
        header.setSpacing(10)
        self.title_lbl = QLabel("모비노기 채집물 자동 판매")
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
        self.btn_theme.setFixedSize(56, 34)
        self.btn_theme.setText("설정")
        self.btn_theme.setToolTip("설정")
        self.btn_theme.clicked.disconnect()
        self.btn_theme.clicked.connect(self.show_settings_popup)
        self.main_layout.addLayout(header)

        self.selected_list = QListWidget()
        self.selected_list.setFixedHeight(96)
        self.selected_list.setObjectName("selectedList")
        self.selected_list.setIconSize(QSize(72, 26))
        self.selected_list.setSpacing(4)
        self.selected_list.setFocusPolicy(Qt.NoFocus)
        self.main_layout.addWidget(self.selected_list)

        self.settings_card = QFrame()
        self.settings_card.setObjectName("settingsCard")
        settings_layout = QVBoxLayout(self.settings_card)
        settings_layout.setContentsMargins(14, 10, 14, 10)
        settings_layout.setSpacing(8)

        self.settings_title_lbl = QLabel("설정")
        self.settings_title_lbl.setObjectName("cardTitle")
        self.settings_title_lbl.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        settings_layout.addWidget(self.settings_title_lbl)

        settings_row = QHBoxLayout()
        settings_row.setSpacing(8)

        self.capture_key_lbl = QLabel("판매 아이템 선택 키")
        self.capture_key_lbl.setObjectName("settingsFieldLabel")
        self.capture_key_lbl.setFont(QFont("Malgun Gothic", 8, QFont.Bold))

        self.capture_key_input = QLineEdit()
        self.capture_key_input.setObjectName("settingsKeyInput")
        self.capture_key_input.setMaxLength(1)
        self.capture_key_input.setFixedSize(52, 34)
        self.capture_key_input.setAlignment(Qt.AlignCenter)
        self.capture_key_input.returnPressed.connect(self.apply_capture_key_from_panel)

        self.btn_capture_key_save = QPushButton("저장")
        self.btn_capture_key_save.setObjectName("settingsSaveButton")
        self.btn_capture_key_save.setFixedHeight(34)
        self.btn_capture_key_save.setCursor(Qt.PointingHandCursor)
        self.btn_capture_key_save.clicked.connect(self.apply_capture_key_from_panel)

        settings_row.addWidget(self.capture_key_lbl)
        settings_row.addStretch(1)
        settings_row.addWidget(self.capture_key_input)
        settings_row.addWidget(self.btn_capture_key_save)
        settings_layout.addLayout(settings_row)

        self.settings_hint_lbl = QLabel("한 글자 키만 설정할 수 있습니다.")
        self.settings_hint_lbl.setObjectName("settingsHintLabel")
        self.settings_hint_lbl.setFont(QFont("Malgun Gothic", 8))
        settings_layout.addWidget(self.settings_hint_lbl)

        self.main_layout.insertWidget(1, self.settings_card)
        self.settings_card.hide()

        self.item_card = QFrame()
        self.item_card.setObjectName("modeCard")
        item_card_layout = QVBoxLayout(self.item_card)
        item_card_layout.setContentsMargins(14, 10, 14, 10)
        item_card_layout.setSpacing(8)

        item_title = QLabel("판매 아이템")
        item_title.setFont(QFont("Malgun Gothic", 9, QFont.Bold))
        item_title.setObjectName("cardTitle")
        item_card_layout.addWidget(item_title)

        action_layout = QHBoxLayout()
        action_layout.setSpacing(8)

        self.btn_select_item = QPushButton("판매 아이템 선택")
        self.btn_select_item.setFixedHeight(34)
        self.btn_select_item.setCursor(Qt.PointingHandCursor)
        self.btn_select_item.clicked.connect(self.start_item_selection)

        self.btn_reset_items = QPushButton("초기화")
        self.btn_reset_items.setFixedHeight(34)
        self.btn_reset_items.setCursor(Qt.PointingHandCursor)
        self.btn_reset_items.clicked.connect(self.clear_templates)

        self.btn_resize_game = QPushButton("창 크기 맞추기")
        self.btn_resize_game.setFixedHeight(34)
        self.btn_resize_game.setCursor(Qt.PointingHandCursor)
        self.btn_resize_game.clicked.connect(self.resize_game_window)

        action_layout.addWidget(self.btn_select_item)
        action_layout.addWidget(self.btn_reset_items)
        action_layout.addWidget(self.btn_resize_game)
        item_card_layout.addLayout(action_layout)

        self.template_count_lbl = QLabel("등록된 판매 아이템: 0개")
        self.template_count_lbl.setObjectName("guideText")
        self.template_count_lbl.setFont(QFont("Malgun Gothic", 8))
        item_card_layout.addWidget(self.template_count_lbl)
        self.main_layout.addWidget(self.item_card)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.btn_start = QPushButton("▶  자동 판매 시작")
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
        self.log_label = QLabel("실행 로그")
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
        self.log_text.setAlignment(Qt.AlignLeft)
        self.log_text.document().setDocumentMargin(0)

        self.btn_clear = QPushButton("지우기", self.log_text)
        self.btn_clear.setFixedWidth(55)
        self.btn_clear.setFixedHeight(22)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self.clear_logs)
        self.main_layout.addWidget(self.log_text)

        guide_layout = QHBoxLayout()
        self.notice_content = QLabel(
            "• 판매 아이템 선택 후 자동 판매 시작\n"
            "• 시작 시 창 크기를 자동으로 한 번 맞추고, 버튼으로 다시 맞출 수 있음\n"
            "• 자동 판매 중에는 화면이 가려지지 않게 유지\n"
            "• 선택 정보는 메모리 캐시라 프로그램 종료 시 사라짐"
        )
        self.notice_content.setObjectName("guideText")
        self.notice_content.setFont(QFont("Malgun Gothic", 8))
        self.notice_content.setAlignment(Qt.AlignLeft | Qt.AlignTop)

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
        self.update_notice_text()
        self.sync_capture_key_panel()
        self.update_template_summary()
        self.refresh_controls()

    def apply_style(self):
        c = self.colors
        self.setStyleSheet(
            f"""
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
            QListWidget#selectedList {{
                background-color: {c['log_bg']};
                color: {c['log_text']};
                border: 1px solid {c['card_border']};
                border-radius: 10px;
                padding: 6px;
                outline: none;
            }}
            QListWidget#selectedList::item {{
                background-color: {c['input_bg']};
                border: 1px solid {c['card_border']};
                border-radius: 8px;
                padding: 6px;
                margin: 1px 0;
            }}
            QListWidget#selectedList::item:selected {{
                background-color: {c['surface']};
                color: {c['text']};
            }}

            QFrame#modeCard {{
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 12px;
            }}
            QFrame#settingsCard {{
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 12px;
            }}
            QLabel#cardTitle {{
                color: {c['accent']};
                border: none;
                background: transparent;
            }}
            QLabel#settingsFieldLabel {{
                color: {c['text']};
                border: none;
                background: transparent;
            }}
            QLabel#settingsHintLabel {{
                color: {c['text_dim']};
                border: none;
                background: transparent;
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

            QPushButton#btnTheme {{
                background-color: {c['card']};
                border: 1px solid {c['card_border']};
                border-radius: 10px;
                font-size: 9pt;
                font-weight: bold;
            }}
            QPushButton#btnTheme:hover {{
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

            QPushButton#btnSelectItem {{
                background-color: {c['accent']};
                color: white;
                border: none;
                border-radius: 8px;
                font-weight: bold;
            }}
            QPushButton#btnSelectItem:hover {{
                background-color: {c['accent_hover']};
            }}
            QPushButton#btnSelectItem:disabled,
            QPushButton#btnTemplateReset:disabled,
            QPushButton#btnWindowResize:disabled,
            QPushButton#settingsSaveButton:disabled {{
                background-color: {c['card']};
                color: {c['text_dim']};
                border: 1px solid {c['card_border']};
            }}

            QPushButton#btnTemplateReset {{
                background-color: {c['danger']};
                color: white;
                border: none;
                border-radius: 8px;
            }}
            QPushButton#btnTemplateReset:hover {{
                background-color: {c['danger_hover']};
            }}

            QPushButton#btnWindowResize {{
                background-color: {c['card']};
                color: {c['text']};
                border: 2px solid #1f5c42;
                border-radius: 8px;
            }}
            QPushButton#btnWindowResize:hover {{
                background-color: {c['surface']};
                border: 2px solid #2f7a58;
            }}

            QPushButton#settingsSaveButton {{
                background-color: {c['accent']};
                color: white;
                border: none;
                border-radius: 8px;
                font-weight: bold;
            }}
            QPushButton#settingsSaveButton:hover {{
                background-color: {c['accent_hover']};
            }}

            QLineEdit#settingsKeyInput {{
                background-color: {c['input_bg']};
                color: {c['text']};
                border: 1px solid {c['input_border']};
                border-radius: 8px;
                padding: 2px 6px;
                font-size: 11pt;
                font-weight: bold;
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
            """
        )

        self.btn_start.setObjectName("btnStart")
        self.btn_stop.setObjectName("btnStop")
        self.btn_guide.setObjectName("btnGuide")
        self.btn_clear.setObjectName("btnClear")
        self.btn_select_item.setObjectName("btnSelectItem")
        self.btn_reset_items.setObjectName("btnTemplateReset")
        self.btn_resize_game.setObjectName("btnWindowResize")
        self.capture_key_lbl.setObjectName("settingsFieldLabel")
        self.capture_key_input.setObjectName("settingsKeyInput")
        self.btn_capture_key_save.setObjectName("settingsSaveButton")
        self.settings_hint_lbl.setObjectName("settingsHintLabel")
        self.log_text.setObjectName("logText")
        self.selected_list.setObjectName("selectedList")

        for widget in [
            self.btn_start,
            self.btn_stop,
            self.btn_guide,
            self.btn_clear,
            self.btn_theme,
            self.btn_select_item,
            self.btn_reset_items,
            self.btn_resize_game,
            self.capture_key_input,
            self.btn_capture_key_save,
            self.log_text,
            self.selected_list,
        ]:
            widget.style().unpolish(widget)
            widget.style().polish(widget)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "btn_clear"):
            self.btn_clear.move(self.log_text.width() - 65, 5)

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        self.colors = DARK_COLORS.copy() if self.is_dark_mode else LIGHT_COLORS.copy()
        self.btn_theme.setText("☀️" if self.is_dark_mode else "🌙")
        self.btn_theme.setToolTip("라이트모드 전환" if self.is_dark_mode else "다크모드 전환")
        self.apply_style()

    def set_theme_mode(self, enabled: bool):
        self.is_dark_mode = enabled
        self.colors = DARK_COLORS.copy() if self.is_dark_mode else LIGHT_COLORS.copy()
        self.apply_style()
        if self.settings_popup and self.settings_popup.isVisible():
            self.settings_popup.refresh_ui()

    def toggle_theme(self):
        self.set_theme_mode(not self.is_dark_mode)

    def toggle_theme(self):
        if self.settings_popup and self.settings_popup.isVisible():
            self.settings_popup.refresh_ui()

    def sync_capture_key_panel(self):
        if hasattr(self, "capture_key_input"):
            self.capture_key_input.setText(format_capture_key(self.capture_key))

    def apply_capture_key_from_panel(self):
        value = normalize_capture_key(self.capture_key_input.text())
        if value is None:
            self.sync_capture_key_panel()
            self.append_log("설정 실패: 판매 아이템 선택 키는 한 글자만 사용할 수 있습니다.")
            return
        self.set_capture_key(value)

    def set_capture_key(self, capture_key: str):
        normalized = normalize_capture_key(capture_key)
        if normalized is None:
            return
        previous_key = self.capture_key
        self.capture_key = normalized
        self.sync_capture_key_panel()
        self.update_notice_text()
        if previous_key != normalized:
            self.append_log(f"설정 변경: 판매 아이템 선택 키 = {format_capture_key(self.capture_key)}")

    def update_notice_text(self):
        display_key = format_capture_key(self.capture_key)
        self.notice_content.setText(
            f"• 판매 아이템 선택 버튼을 누른 후 인벤토리를 열고 \n아이템에 마우스를 올리고 {display_key} 키를 누르세요.\n"
            "• 시작 시 창 크기를 자동으로 한 번 맞추고, 필요하면 버튼으로 다시 맞출 수 있습니다.\n"
            "• 자동 판매는 한 품목을 팔 때마다 현재 인벤토리를 다시 검사합니다."
        )

    def show_settings_popup(self):
        self.is_settings_open = not self.is_settings_open
        self.settings_card.setVisible(self.is_settings_open)
        self.btn_theme.setText("설정 ▲" if self.is_settings_open else "설정")
        self.setFixedSize(400, self.expanded_window_height if self.is_settings_open else self.base_window_height)
        self.sync_capture_key_panel()

    def update_notice_text(self):
        display_key = format_capture_key(self.capture_key)
        self.notice_content.setText(
            f"• 판매 아이템 선택 후 인벤토리를 열고 아이템 위에 커서를 두고\n"
            f"  {display_key} 키를누르세요.\n"
            "• 판매 아이템 선택 모드 종료는 Esc 입니다.\n"
            "• 여러 품목을 판매하고 싶다면 여러번 반복하여 등록하면 됩니다.\n"
            "• 이미지 스캔 방식이라 창이 가려지면 작동하지 않을 수 있습니다."
        )

    def refresh_controls(self):
        can_edit_templates = not self.is_selecting and not self.is_auto_running
        has_templates = bool(self.selected_templates)
        self.btn_theme.setEnabled(can_edit_templates)
        self.btn_select_item.setEnabled(can_edit_templates)
        self.btn_reset_items.setEnabled(can_edit_templates and has_templates)
        self.btn_resize_game.setEnabled(can_edit_templates)
        self.capture_key_input.setEnabled(can_edit_templates)
        self.btn_capture_key_save.setEnabled(can_edit_templates)
        self.btn_start.setEnabled(can_edit_templates and has_templates)
        self.btn_stop.setEnabled(self.is_auto_running)

    def update_template_summary(self):
        self.selected_list.clear()
        if not self.selected_templates:
            placeholder = QListWidgetItem("선택된 판매 아이템이 없습니다.")
            placeholder.setFlags(Qt.NoItemFlags)
            self.selected_list.addItem(placeholder)
        else:
            for index, template in enumerate(self.selected_templates, start=1):
                item = QListWidgetItem(
                    pil_image_to_icon(template.template_image),
                    f"{index}. {template.label}",
                )
                item.setSizeHint(QSize(0, 42))
                self.selected_list.addItem(item)
        self.template_count_lbl.setText(f"등록된 판매 아이템: {len(self.selected_templates)}개")

    def start_item_selection(self):
        if self.is_selecting or self.is_auto_running:
            return
        self.is_selecting = True
        self.refresh_controls()
        self.append_log("\n=== 판매 아이템 선택 대기 중 ===")
        self.append_log("인벤토리를 열고 판매할 아이템 위에 마우스를 올린 상태에서 설정된 선택 키를 누르세요. 취소는 Esc입니다.")

        self.append_log(
            f"현재 선택 키: {format_capture_key(self.capture_key)}. 인벤토리를 열고 아이템 위에서 해당 키를 누르세요."
        )
        self.selection_worker = ItemSelectionWorker(len(self.selected_templates) + 1, self.capture_key)
        self.selection_worker.log_signal.connect(self.append_log)
        self.selection_worker.template_selected_signal.connect(self.on_template_selected)
        self.selection_worker.finished_signal.connect(self.on_selection_finished)

        self.selection_thread = threading.Thread(target=self.selection_worker.run, daemon=True)
        self.selection_thread.start()

    def clear_templates(self):
        if self.is_selecting or self.is_auto_running or not self.selected_templates:
            return
        self.selected_templates.clear()
        self.update_template_summary()
        self.refresh_controls()
        self.append_log("🧹 판매 아이템 캐시 초기화")

    def resize_game_window(self, checked=False, auto=False):
        try:
            hwnd = win32gui.FindWindow(None, WINDOW_TITLE)
            if not hwnd:
                self.append_log("⚠️ 창 크기 맞추기 실패: 게임 창을 찾지 못했습니다.")
                return
            client_width, client_height = resize_client_area_for_hwnd(hwnd)
            prefix = "자동" if auto else "수동"
            self.append_log(
                f"{prefix} 창 크기 맞춤 완료: {client_width} x {client_height}\n매크로 사용중에는 클라이언트 창 크기를 변경하지 말아주세요"
            )
        except Exception as exc:
            prefix = "자동" if auto else "수동"
            self.append_log(f"⚠️ {prefix} 창 크기 맞추기 실패: {exc}")

    def start_macro(self):
        if self.is_selecting or self.is_auto_running:
            return
        if not self.selected_templates:
            self.append_log("❌ 판매 아이템이 선택되지 않았습니다.")
            return

        self.is_auto_running = True
        self.refresh_controls()

        templates = [
            ItemTemplate(
                label=template.label,
                template_image=template.template_image.copy(),
                source_row=template.source_row,
                source_col=template.source_col,
                created_at=template.created_at,
            )
            for template in self.selected_templates
        ]

        self.auto_worker = AutoSellWorker(templates)
        self.auto_worker.log_signal.connect(self.append_log)
        self.auto_worker.countdown_signal.connect(self.update_countdown)
        self.auto_worker.finished_signal.connect(self.on_auto_finished)

        self.auto_thread = threading.Thread(target=self.auto_worker.run, daemon=True)
        self.auto_thread.start()

    def stop_macro(self):
        if self.auto_worker:
            self.auto_worker.stop()
        self.btn_stop.setEnabled(False)

    def clear_logs(self):
        self.log_text.clear()
        self.append_log("── 로그 초기화 ──")

    def show_guide(self):
        notice = NoticeDialog(self.colors, self.capture_key, self)
        notice.exec_()

    @pyqtSlot(object)
    def on_template_selected(self, template_obj):
        self.selected_templates.append(template_obj)
        self.update_template_summary()

    @pyqtSlot()
    def on_selection_finished(self):
        self.is_selecting = False
        self.selection_worker = None
        self.refresh_controls()

    @pyqtSlot()
    def on_auto_finished(self):
        self.is_auto_running = False
        self.auto_worker = None
        self.countdown_lbl.setText("")
        self.refresh_controls()

    @pyqtSlot(str)
    def append_log(self, text):
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        if self.log_text.toPlainText():
            cursor.insertBlock()
        block_format = QTextBlockFormat()
        block_format.setAlignment(Qt.AlignLeft)
        cursor.setBlockFormat(block_format)
        cursor.insertText(text)
        self.log_text.setTextCursor(cursor)
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    @pyqtSlot(int)
    def update_countdown(self, seconds):
        if seconds >= 0:
            self.countdown_lbl.setText(f"다음 실행까지: {seconds}초")
        else:
            self.countdown_lbl.setText("")

    def closeEvent(self, event):
        if self.settings_popup:
            self.settings_popup.close()
        if self.auto_worker:
            self.auto_worker.stop()
        if self.selection_worker:
            self.selection_worker.stop()
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()

    notice = NoticeDialog(window.colors, window.capture_key, window)
    notice.exec_()
    sys.exit(app.exec_())
