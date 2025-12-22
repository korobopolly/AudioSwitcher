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

from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize, GUID
from pycaw.pycaw import AudioUtilities, IMMDeviceEnumerator
from pycaw.constants import CLSID_MMDeviceEnumerator
import comtypes
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item


# Config file path (same directory as script)
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'audio_switcher_config.json')


def kill_existing_instance():
    """Kill any existing instance of this application."""
    current_pid = os.getpid()

    # Kill pythonw.exe running audio_switcher.py
    try:
        result = subprocess.run(
            ['wmic', 'process', 'where',
             "name='pythonw.exe' and commandline like '%audio_switcher.py%'",
             'get', 'processid'],
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

    # Kill AudioSwitcher.exe
    try:
        result = subprocess.run(
            ['wmic', 'process', 'where', "name='AudioSwitcher.exe'", 'get', 'processid'],
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

    def set_favorite(self, slot, device_id):
        """Set a device as favorite."""
        self._config[f'favorite{slot}'] = device_id
        save_config(self._config)

    def create_icon_image(self):
        """Create a switch/swap icon."""
        size = 64
        image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)

        # Two horizontal arrows (swap symbol)
        draw.line([(12, 20), (52, 20)], fill='white', width=4)
        draw.polygon([(52, 20), (42, 12), (42, 28)], fill='white')

        draw.line([(12, 44), (52, 44)], fill='white', width=4)
        draw.polygon([(12, 44), (22, 36), (22, 52)], fill='white')

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
        return callback

    def _make_favorite_callback(self, slot, device_id):
        """Create callback for setting favorite."""
        def callback(icon, item):
            self.set_favorite(slot, device_id)
        return callback

    def _on_click(self, icon, item):
        """Handle left-click on icon."""
        self.toggle_favorites()

    def _on_refresh(self, icon, item):
        """Refresh device list."""
        self._refresh_devices()
        icon.update_menu()

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
