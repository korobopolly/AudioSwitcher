"""
Audio Output Switcher - System Tray Application
Allows switching between audio output devices from the system tray.
Left-click: Toggle between two favorite devices
Right-click: Menu to select device or set favorites
"""

import warnings
warnings.filterwarnings('ignore')

import sys
import os
import json
import subprocess
import ctypes
from ctypes import wintypes
import winreg
import urllib.request
import threading
import tempfile

# 버전 정보
GITHUB_REPO = "korobopolly/AudioSwitcher"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def get_version():
    """Git tag 또는 GitHub 릴리즈에서 현재 버전을 가져옵니다."""
    # 1. 환경변수에서 버전 확인 (빌드 시 설정)
    version = os.environ.get('APP_VERSION')
    if version:
        return version.lstrip('v')

    # 2. 버전 파일에서 확인
    version_file = os.path.join(os.path.dirname(__file__), 'VERSION')
    if os.path.exists(version_file):
        try:
            with open(version_file, 'r') as f:
                return f.read().strip().lstrip('v')
        except Exception:
            pass

    # 3. Git tag에서 확인 (개발 환경)
    try:
        result = subprocess.run(
            ['git', 'describe', '--tags', '--abbrev=0'],
            capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        if result.returncode == 0:
            return result.stdout.strip().lstrip('v')
    except Exception:
        pass

    return "0.0.0"  # 기본값


VERSION = get_version()

from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize, GUID
from pycaw.pycaw import AudioUtilities, IMMDeviceEnumerator
from pycaw.constants import CLSID_MMDeviceEnumerator
import comtypes
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item


# Config file path
# PyInstaller로 패키징된 경우 실행 파일 위치 사용, 아니면 스크립트 위치 사용
def get_config_path():
    if getattr(sys, 'frozen', False):
        # PyInstaller로 패키징된 .exe 파일인 경우
        base_dir = os.path.dirname(sys.executable)
    else:
        # 일반 Python 스크립트로 실행되는 경우
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, 'audio_switcher_config.json')

CONFIG_FILE = get_config_path()


def is_windows_light_theme():
    """윈도우 테마가 라이트 모드인지 확인합니다."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        )
        # SystemUsesLightTheme: 시스템 트레이 영역 테마
        value, _ = winreg.QueryValueEx(key, "SystemUsesLightTheme")
        winreg.CloseKey(key)
        return value == 1  # 1 = 라이트 모드, 0 = 다크 모드
    except Exception:
        return False  # 기본값: 다크 모드로 가정


def http_request(url, timeout=10):
    """HTTP GET 요청을 수행합니다. 응답 데이터 또는 None 반환"""
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'AudioSwitcher'})
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return response.read()
    except Exception:
        return None


def compare_versions(v1, v2):
    """버전 비교. v1 > v2이면 1, v1 < v2이면 -1, 같으면 0 반환"""
    try:
        parts1 = [int(x) for x in v1.split('.')]
        parts2 = [int(x) for x in v2.split('.')]
        max_len = max(len(parts1), len(parts2))
        parts1.extend([0] * (max_len - len(parts1)))
        parts2.extend([0] * (max_len - len(parts2)))
        for p1, p2 in zip(parts1, parts2):
            if p1 > p2:
                return 1
            elif p1 < p2:
                return -1
        return 0
    except Exception:
        return 0


def check_for_updates():
    """GitHub에서 최신 버전을 확인합니다. (버전, 다운로드URL, 릴리즈노트) 또는 None 반환"""
    try:
        data = http_request(GITHUB_API_URL)
        if not data:
            return None
        data = json.loads(data.decode('utf-8'))

        latest_version = data.get('tag_name', '').lstrip('v')
        release_notes = data.get('body', '')

        # exe 파일 다운로드 URL 찾기
        download_url = None
        for asset in data.get('assets', []):
            if asset['name'].endswith('.exe'):
                download_url = asset['browser_download_url']
                break

        if not download_url:
            return None

        # 버전 비교 (현재 버전보다 새로운지)
        if compare_versions(latest_version, VERSION) > 0:
            return (latest_version, download_url, release_notes)

        return None  # 최신 버전 사용 중
    except Exception:
        return None


def download_and_apply_update(download_url, callback=None):
    """업데이트를 다운로드하고 적용합니다."""
    try:
        # 스크립트 모드에서는 업데이트 불가
        if not getattr(sys, 'frozen', False):
            if callback:
                callback(False, "스크립트 모드에서는 자동 업데이트가 불가능합니다.")
            return

        current_exe = sys.executable
        temp_dir = tempfile.gettempdir()
        new_exe = os.path.join(temp_dir, "AudioSwitcher_new.exe")
        updater_bat = os.path.join(temp_dir, "update_audioswitcher.bat")

        # 다운로드
        data = http_request(download_url, timeout=60)
        if not data:
            if callback:
                callback(False, "다운로드 실패")
            return

        with open(new_exe, 'wb') as f:
            f.write(data)

        # 업데이트 배치 스크립트 생성
        bat_content = f'''@echo off
REM Wait for the app to close
timeout /t 2 /nobreak >nul
REM Replace the old exe with the new one
copy /Y "{new_exe}" "{current_exe}"
REM Start the updated app
start "" "{current_exe}"
REM Clean up
del "{new_exe}"
del "%~f0"
'''
        with open(updater_bat, 'w') as f:
            f.write(bat_content)

        subprocess.Popen(['cmd', '/c', updater_bat], creationflags=subprocess.CREATE_NO_WINDOW)

        if callback:
            callback(True, "업데이트 적용 중... 앱이 재시작됩니다.")
        os._exit(0)

    except Exception as e:
        if callback:
            callback(False, f"업데이트 실패: {str(e)}")


def kill_process_by_name(process_filter, current_pid):
    """wmic으로 프로세스를 찾아 종료합니다."""
    try:
        result = subprocess.run(
            ['wmic', 'process', 'where', process_filter, 'get', 'processid'],
            capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        for line in result.stdout.strip().split('\n'):
            line = line.strip()
            if line.isdigit():
                pid = int(line)
                if pid != current_pid:
                    subprocess.run(['taskkill', '/f', '/pid', str(pid)],
                                   capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception:
        pass


def kill_existing_instance():
    """Kill any existing instance of this application."""
    current_pid = os.getpid()
    kill_process_by_name("name='pythonw.exe' and commandline like '%audio_switcher.py%'", current_pid)
    kill_process_by_name("name='AudioSwitcher.exe'", current_pid)


def load_config():
    """Load configuration from file."""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {'favorite1': None, 'favorite2': None}


def save_config(config):
    """Save configuration to file."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


class PolicyConfigClient:
    """Interface to set default audio endpoint."""

    def __init__(self):
        self._policy_config = None
        self._init_policy_config()

    def _init_policy_config(self):
        try:
            IID_IPolicyConfig = GUID("{F8679F50-850A-41CF-9C72-430F290290C8}")
            CLSID_PolicyConfigClient = GUID("{870AF99C-171D-4F9E-AF0D-E63DF40C2BC9}")

            class IPolicyConfig(comtypes.IUnknown):
                _iid_ = IID_IPolicyConfig
                _methods_ = [
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused1'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused2'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused3'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused4'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused5'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused6'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused7'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused8'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused9'),
                    comtypes.COMMETHOD([], comtypes.HRESULT, 'Unused10'),
                    comtypes.COMMETHOD(
                        [], comtypes.HRESULT, 'SetDefaultEndpoint',
                        (['in'], comtypes.c_wchar_p, 'deviceId'),
                        (['in'], comtypes.c_uint, 'role')
                    ),
                ]

            self._policy_config = comtypes.CoCreateInstance(
                CLSID_PolicyConfigClient,
                IPolicyConfig,
                CLSCTX_ALL
            )
        except Exception as e:
            self._policy_config = None

    def set_default_endpoint(self, device_id: str):
        """Set default audio endpoint for all roles."""
        if self._policy_config:
            for role in range(3):
                try:
                    self._policy_config.SetDefaultEndpoint(device_id, role)
                except Exception:
                    pass


class AudioSwitcher:
    """Main application class for audio switching."""

    def __init__(self):
        self.icon = None
        self.policy_client = None
        self._running = True
        self._devices = []
        self._config = load_config()
        self._refresh_devices()

    def _refresh_devices(self):
        """Refresh the list of audio devices."""
        self._devices = []
        try:
            all_devices = AudioUtilities.GetAllDevices()
            for dev in all_devices:
                if dev.state.name == 'Active' and dev.id:
                    if hasattr(dev, 'flow') and dev.flow is not None:
                        if dev.flow.value != 0:
                            continue

                    name = dev.FriendlyName or ""
                    skip_patterns = ['Microphone', 'Mic', 'Input', 'Line In', 'Rear Green In',
                                     'Rear Blue In', 'Front Green In', 'Front Pink In',
                                     'Rear Pink In']
                    if any(pattern.lower() in name.lower() for pattern in skip_patterns):
                        continue

                    self._devices.append({
                        'id': dev.id,
                        'name': name or f"Device {len(self._devices) + 1}"
                    })
        except Exception:
            pass

    def get_default_device_id(self):
        """Get current default audio output device ID."""
        try:
            deviceEnumerator = comtypes.CoCreateInstance(
                CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                CLSCTX_ALL
            )
            default_device = deviceEnumerator.GetDefaultAudioEndpoint(0, 0)
            return default_device.GetId()
        except Exception:
            return None

    def set_default_device(self, device_id: str):
        """Set default audio output device."""
        if self.policy_client is None:
            self.policy_client = PolicyConfigClient()
        self.policy_client.set_default_endpoint(device_id)

    def get_device_name(self, device_id):
        """Get device name by ID."""
        for dev in self._devices:
            if dev['id'] == device_id:
                return dev['name']
        return None

    def get_current_favorite_number(self):
        """현재 기본 장치가 즐겨찾기 몇 번인지 반환합니다. (1, 2, 또는 None)"""
        current = self.get_default_device_id()
        fav1 = self._config.get('favorite1')
        fav2 = self._config.get('favorite2')

        if current == fav1:
            return 1
        elif current == fav2:
            return 2
        return None

    def update_icon(self):
        """현재 상태에 맞게 아이콘을 업데이트합니다."""
        if self.icon:
            self.icon.icon = self.create_icon_image()

    def toggle_favorites(self):
        """Toggle between two favorite devices."""
        fav1 = self._config.get('favorite1')
        fav2 = self._config.get('favorite2')

        if not fav1 or not fav2:
            return  # Favorites not set

        current = self.get_default_device_id()

        if current == fav1:
            self.set_default_device(fav2)
        else:
            self.set_default_device(fav1)

        # 아이콘 업데이트
        self.update_icon()

    def set_favorite(self, slot, device_id):
        """Set a device as favorite."""
        self._config[f'favorite{slot}'] = device_id
        save_config(self._config)

    def create_icon_image(self):
        """큰 숫자 중심의 시인성 좋은 아이콘을 생성합니다."""
        size = 64
        image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)

        # 테마에 따른 색상 결정
        if is_windows_light_theme():
            color = (32, 32, 32, 255)  # 검정색
        else:
            color = (255, 255, 255, 255)  # 흰색

        fav_num = self.get_current_favorite_number()

        # 상단에 작은 스피커 아이콘
        # 스피커 몸통
        draw.polygon([(4, 8), (4, 20), (12, 20), (12, 8)], fill=color)
        # 스피커 콘
        draw.polygon([(12, 8), (12, 20), (22, 26), (22, 2)], fill=color)
        # 음파
        draw.arc([(20, 4), (32, 24)], start=-60, end=60, fill=color, width=2)
        draw.arc([(26, 0), (40, 28)], start=-60, end=60, fill=color, width=2)

        # 큰 숫자 표시 (중앙~하단)
        if fav_num:
            if fav_num == 1:
                # 큰 "1" 그리기
                draw.line([(32, 24), (32, 60)], fill=color, width=6)
                draw.line([(22, 32), (32, 24)], fill=color, width=4)
                draw.line([(22, 60), (42, 60)], fill=color, width=4)
            else:
                # 큰 "2" 그리기
                draw.arc([(14, 22), (50, 46)], start=180, end=10, fill=color, width=5)
                draw.line([(48, 36), (14, 58)], fill=color, width=5)
                draw.line([(14, 58), (50, 58)], fill=color, width=5)
        else:
            # 즐겨찾기 미설정 시 "?" 표시
            draw.arc([(18, 24), (46, 48)], start=180, end=0, fill=color, width=4)
            draw.line([(46, 36), (32, 50)], fill=color, width=4)
            draw.ellipse([(28, 56), (36, 62)], fill=color)

        return image

    def _is_default(self, device_id):
        """Check if device is the default."""
        def check(item):
            return self.get_default_device_id() == device_id
        return check

    def _is_favorite(self, slot, device_id):
        """Check if device is the favorite for given slot."""
        def check(item):
            return self._config.get(f'favorite{slot}') == device_id
        return check

    def _make_select_callback(self, device_id):
        """Create callback for device selection."""
        def callback(icon, item):
            self.set_default_device(device_id)
            self.update_icon()
        return callback

    def _make_favorite_callback(self, slot, device_id):
        """Create callback for setting favorite."""
        def callback(icon, item):
            self.set_favorite(slot, device_id)
            self.update_icon()
        return callback

    def _on_click(self, icon, item):
        """Handle left-click on icon."""
        self.toggle_favorites()

    def _on_refresh(self, icon, item):
        """Refresh device list."""
        self._refresh_devices()
        icon.update_menu()

    def _on_check_update(self, icon, item):
        """업데이트 확인 및 적용."""
        def check_and_update():
            update_info = check_for_updates()
            if update_info:
                version, url, notes = update_info
                # 업데이트 다운로드 및 적용
                download_and_apply_update(url)
            else:
                # 최신 버전 사용 중 - 알림 표시
                try:
                    ctypes.windll.user32.MessageBoxW(
                        0,
                        f"현재 최신 버전(v{VERSION})을 사용 중입니다.",
                        "업데이트 확인",
                        0x40  # MB_ICONINFORMATION
                    )
                except Exception:
                    pass

        # 백그라운드에서 실행
        thread = threading.Thread(target=check_and_update, daemon=True)
        thread.start()

    def _create_favorite_submenu(self, slot):
        """Create submenu for setting favorite."""
        menu_items = []
        for device in self._devices:
            menu_items.append(
                item(
                    device['name'],
                    self._make_favorite_callback(slot, device['id']),
                    checked=self._is_favorite(slot, device['id'])
                )
            )
        return pystray.Menu(*menu_items)

    def _get_favorite_label(self, slot):
        """Get dynamic label for favorite slot."""
        def get_label(item):
            name = self.get_device_name(self._config.get(f'favorite{slot}')) or '(Not set)'
            return f'즐찾 [{slot}] {name}'
        return get_label

    def create_menu(self):
        """Create menu with audio devices."""
        menu_items = []

        # Toggle info with dynamic labels
        menu_items.append(item(self._get_favorite_label(1), self._create_favorite_submenu(1)))
        menu_items.append(item(self._get_favorite_label(2), self._create_favorite_submenu(2)))

        menu_items.append(pystray.Menu.SEPARATOR)

        # All devices
        for device in self._devices:
            menu_items.append(
                item(
                    device['name'],
                    self._make_select_callback(device['id']),
                    checked=self._is_default(device['id'])
                )
            )

        menu_items.append(pystray.Menu.SEPARATOR)
        menu_items.append(item('Refresh', self._on_refresh))
        menu_items.append(item(f'Update (v{VERSION})', self._on_check_update))
        menu_items.append(item('Exit', lambda icon, item: icon.stop()))

        return menu_items

    def run(self):
        """Run the application."""
        CoInitialize()
        try:
            image = self.create_icon_image()
            self.icon = pystray.Icon(
                "audio_switcher",
                image,
                "Audio Switcher (Left-click to toggle)",
                menu=pystray.Menu(
                    item('Toggle', self._on_click, default=True, visible=False),
                    *self.create_menu()
                )
            )
            self.icon.run()
        finally:
            CoUninitialize()


def main():
    """Entry point."""
    kill_existing_instance()
    app = AudioSwitcher()
    app.run()


if __name__ == "__main__":
    main()
