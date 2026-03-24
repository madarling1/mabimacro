"""
타임아웃 복구 로직 테스트 스크립트
- recovery_reset: Esc 반복 → Process.png 감지될 때까지
- 타임아웃 감지: ReceiveAll 30초 대기 후 복구 트리거

사용법:
  1. 게임 창(마비노기 모바일)을 열어둔 상태에서 실행
  2. 콘솔에서 로그를 확인하면서 동작 검증
  3. 테스트 모드를 선택해서 개별 테스트 가능
"""

import sys
import os
import time
import cv2
import numpy as np
import ctypes
import win32gui
import win32ui
import win32con
import win32api

# DPI 설정
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    ctypes.windll.user32.SetProcessDPIAware()


# ═══════════════════════════════════════════════
#  타임아웃 상수
# ═══════════════════════════════════════════════
RECEIVE_TIMEOUT = 30   # ReceiveAll 탐색 최대 대기(초)
CRAFT_TIMEOUT = 30     # 재료 탐색 최대 대기(초)
RECOVERY_MAX_ESC = 10  # 복구 시 Esc 최대 횟수


class TimeoutRecoveryTester:
    """타임아웃 복구 로직만 가져온 테스트 클래스"""

    GAME_TITLE = "마비노기 모바일"
    VK_K = 0x4B
    threshold = 0.9

    def __init__(self):
        self.is_running = True
        if getattr(sys, 'frozen', False):
            self.base_path = sys._MEIPASS
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))

    def log(self, msg):
        """콘솔 로그"""
        timestamp = time.strftime("%H:%M:%S")
        print(f"[{timestamp}] {msg}")

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
        except Exception as e:
            self.log(f"⚠️ 스크린샷 실패: {e}")
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
        full_path = os.path.join(self.base_path, image_name)
        if not os.path.exists(full_path):
            self.log(f"⚠️ 이미지 파일 없음: {image_name}")
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
                center = (max_loc[0] + tw // 2, max_loc[1] + th // 2)
                self.log(f"  📍 {image_name} 감지 (유사도: {max_val:.3f}, 위치: {center})")
                return center
            else:
                self.log(f"  ❌ {image_name} 미감지 (최대 유사도: {max_val:.3f} < {threshold})")
        except Exception as e:
            self.log(f"⚠️ 이미지 매칭 오류: {e}")
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

    def find_and_click(self, hwnd, image_name, threshold=0.7, wait_time=5):
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

    # ═══════════════════════════════════════════════
    #  🔧 핵심: 타임아웃 복구 메서드
    # ═══════════════════════════════════════════════
    def recovery_reset(self, parent_hwnd, child_hwnd):
        """타임아웃 복구: Esc 반복 → Process.png 감지될 때까지"""
        self.log("🔧 ═══ 타임아웃 복구 시작 ═══")
        for attempt in range(RECOVERY_MAX_ESC):
            self.send_key_perfect(child_hwnd, win32con.VK_ESCAPE, 0x01, 0.05)
            self.log(f"⌨️ Esc ({attempt + 1}/{RECOVERY_MAX_ESC}회)")
            self.interruptible_sleep(1.5)

            screen = self.get_window_screenshot(parent_hwnd)
            if self.find_image_pos(screen, 'Process.png', threshold=0.7):
                self.log("✅ Process.png 감지 — 복구 완료!")
                self.log("🔧 ═══ 타임아웃 복구 종료 ═══")
                return True

        self.log("⚠️ Process.png 미감지 — 복구 실패 (최대 Esc 횟수 초과)")
        self.log("🔧 ═══ 타임아웃 복구 종료 ═══")
        return False

    # ═══════════════════════════════════════════════
    #  테스트 1: Process.png 감지 테스트 (Esc 없이)
    # ═══════════════════════════════════════════════
    def test_process_detection(self, parent_hwnd):
        """현재 화면에서 Process.png가 보이는지만 확인"""
        self.log("\n═══ 테스트 1: Process.png 감지 테스트 ═══")
        screen = self.get_window_screenshot(parent_hwnd)
        if screen is None:
            self.log("❌ 스크린샷 실패")
            return

        pos = self.find_image_pos(screen, 'Process.png', threshold=0.7)
        if pos:
            self.log(f"✅ Process.png 감지됨 → 위치: {pos}")
        else:
            self.log("❌ Process.png 미감지 (창이 열려있거나 다른 화면일 수 있음)")

        # 추가: 다른 이미지들도 감지해보기
        self.log("\n--- 추가 이미지 감지 ---")
        test_images = [
            'ReceiveAll.png', 'Processing1.png', 'Processing2.png',
            'Error1_NoIngre.png', 'Error2_FullOrder.png',
            'IronIngot.png', 'SteelIngot.png', 'CopperIngot.png',
            'Wood.png', 'Leather.png', 'Fabric.png',
        ]
        for img in test_images:
            self.find_image_pos(screen, img, threshold=0.7)

    # ═══════════════════════════════════════════════
    #  테스트 2: recovery_reset 실행 테스트
    # ═══════════════════════════════════════════════
    def test_recovery_reset(self, parent_hwnd, child_hwnd):
        """Esc 반복 → Process.png 나올 때까지 복구 테스트"""
        self.log("\n═══ 테스트 2: recovery_reset 실행 테스트 ═══")
        self.log("⚠️ 3초 후 Esc 키가 전송됩니다. 게임 창에서 아무 창이나 열어두세요.")
        self.interruptible_sleep(3.0)

        result = self.recovery_reset(parent_hwnd, child_hwnd)
        if result:
            self.log("✅ 복구 성공! Process.png가 감지된 상태입니다.")
        else:
            self.log("❌ 복구 실패. Process.png를 감지할 수 없었습니다.")

    # ═══════════════════════════════════════════════
    #  테스트 3: 타임아웃 → 복구 → K키 흐름 테스트
    # ═══════════════════════════════════════════════
    def test_timeout_and_recovery_flow(self, parent_hwnd, child_hwnd):
        """ReceiveAll 탐색 타임아웃 → 복구 → K키 → 제작대 목록 진입"""
        self.log("\n═══ 테스트 3: 타임아웃 → 복구 → K키 전체 흐름 ═══")
        self.log(f"⏳ ReceiveAll.png을 {RECEIVE_TIMEOUT}초 동안 탐색합니다...")
        self.log("   (없으면 타임아웃 → 복구 → K키 흐름 테스트)")
        self.log("   3초 후 시작...")
        self.interruptible_sleep(3.0)

        # 시뮬레이션: ReceiveAll 탐색 (짧은 타임아웃으로 테스트)
        TEST_TIMEOUT = 10  # 테스트용 짧은 타임아웃
        self.log(f"⏳ ReceiveAll 탐색 시작 (테스트 타임아웃: {TEST_TIMEOUT}초)...")

        start_time = time.time()
        found = False
        while time.time() - start_time < TEST_TIMEOUT:
            self.check_running()
            screen = self.get_window_screenshot(parent_hwnd)
            pos = self.find_image_pos(screen, 'ReceiveAll.png', threshold=self.threshold)
            if pos:
                self.log("✅ ReceiveAll.png 감지됨! (타임아웃 미발생)")
                found = True
                break
            elapsed = int(time.time() - start_time)
            self.log(f"  ⏳ 대기 중... ({elapsed}/{TEST_TIMEOUT}초)")
            self.interruptible_sleep(2.0)

        if not found:
            self.log(f"\n⏰ 타임아웃 발생! ({TEST_TIMEOUT}초 경과)")
            self.log("🔧 복구 흐름 시작...")

            # Step 1: recovery_reset (Esc → Process.png 감지)
            recovered = self.recovery_reset(parent_hwnd, child_hwnd)

            if recovered:
                # Step 2: K키 입력
                self.log("\n--- K키 입력 (제작대 목록 열기) ---")
                self.interruptible_sleep(1.5)
                self.send_key_perfect(child_hwnd, self.VK_K, 0x25, 0.05)
                self.log("⌨️ K키 전송 완료")
                self.interruptible_sleep(1.5)

                # Step 3: Processing1/2.png 탐색
                self.log("\n--- Processing1/2.png 탐색 (제작대 목록) ---")
                for proc_img in ['Processing1.png', 'Processing2.png']:
                    screen = self.get_window_screenshot(parent_hwnd)
                    pos = self.find_image_pos(screen, proc_img, threshold=0.7)
                    if pos:
                        self.log(f"✅ {proc_img} 감지! 클릭합니다.")
                        self.background_click_pro(parent_hwnd, pos[0], pos[1])
                        break
                else:
                    self.log("⚠️ Processing1/2.png 모두 미감지")

                self.log("\n✅ 타임아웃 → 복구 → K키 전체 흐름 완료!")
            else:
                self.log("\n❌ 복구 실패로 K키 흐름 진행 불가")


def find_game_window():
    """게임 창 핸들 찾기"""
    title = "마비노기 모바일"
    parent_hwnd = win32gui.FindWindow(None, title)
    if not parent_hwnd:
        print(f"❌ '{title}' 창을 찾을 수 없습니다.")
        print("   게임을 먼저 실행해주세요.")
        return None, None
    child_hwnd = win32gui.FindWindowEx(parent_hwnd, None, None, None) or parent_hwnd
    print(f"✅ 게임 창 발견 (parent: {parent_hwnd}, child: {child_hwnd})")
    return parent_hwnd, child_hwnd


def main():
    print("=" * 50)
    print("  타임아웃 복구 로직 테스트")
    print("=" * 50)

    parent_hwnd, child_hwnd = find_game_window()
    if not parent_hwnd:
        return

    tester = TimeoutRecoveryTester()

    print("\n테스트 메뉴:")
    print("  1. Process.png 감지 테스트 (Esc 없이 화면만 스캔)")
    print("  2. recovery_reset 테스트 (Esc 반복 → Process.png 감지)")
    print("  3. 전체 흐름 테스트 (타임아웃 → 복구 → K키)")
    print("  q. 종료")

    while True:
        choice = input("\n선택> ").strip()

        if choice == '1':
            tester.test_process_detection(parent_hwnd)
        elif choice == '2':
            tester.test_recovery_reset(parent_hwnd, child_hwnd)
        elif choice == '3':
            tester.test_timeout_and_recovery_flow(parent_hwnd, child_hwnd)
        elif choice.lower() == 'q':
            print("종료합니다.")
            break
        else:
            print("1, 2, 3, q 중 선택해주세요.")


if __name__ == '__main__':
    main()
