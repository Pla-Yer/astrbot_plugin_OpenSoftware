import os
import subprocess
import sys
import winreg
from pathlib import Path
from typing import Optional, List, Dict, Tuple
import warnings

# Try to import pywin32 for .lnk parsing
try:
    import win32com.client
    import pythoncom
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False
    warnings.warn("pywin32 is not installed. LnkResolver will not work.")

from astrbot.api.event import filter, AstrMessageEvent, MessageEventResult
from astrbot.api.star import Context, Star, register
from astrbot.api import logger


class RegistrySearcher:
    """Search for installed applications in Windows Registry."""

    # Common registry paths where applications are registered
    REG_PATHS = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\Applications"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Classes\Applications"),
    ]

    @classmethod
    def search_app(cls, app_name: str) -> Optional[str]:
        """
        Search for an application by name in Windows Registry.

        Args:
            app_name: Name of the application (e.g., "chrome.exe", "notepad.exe")

        Returns:
            Full path to the executable if found, None otherwise.
        """
        # Try with .exe extension if not already present
        if not app_name.lower().endswith('.exe'):
            app_name_exe = app_name + '.exe'
        else:
            app_name_exe = app_name
            app_name = app_name[:-4]  # Remove .exe for searching without extension

        for hive, reg_path in cls.REG_PATHS:
            try:
                # Try with .exe extension
                path = cls._try_registry_path(hive, reg_path, app_name_exe)
                if path and os.path.exists(path):
                    return path

                # Try without .exe extension
                path = cls._try_registry_path(hive, reg_path, app_name)
                if path and os.path.exists(path):
                    return path

            except (OSError, WindowsError, FileNotFoundError):
                continue

        return None

    @staticmethod
    def _try_registry_path(hive, reg_path, key_name: str) -> Optional[str]:
        """Try to read a registry path and return the executable path."""
        try:
            with winreg.OpenKey(hive, reg_path) as root_key:
                try:
                    with winreg.OpenKey(root_key, key_name) as app_key:
                        # Read the default value (empty string) which usually contains the path
                        path, _ = winreg.QueryValueEx(app_key, "")
                        return path
                except (OSError, WindowsError, FileNotFoundError):
                    # Key not found, try to enumerate subkeys
                    try:
                        i = 0
                        while True:
                            subkey_name = winreg.EnumKey(root_key, i)
                            if subkey_name.lower() == key_name.lower():
                                with winreg.OpenKey(root_key, subkey_name) as app_key:
                                    path, _ = winreg.QueryValueEx(app_key, "")
                                    return path
                            i += 1
                    except OSError:
                        # No more subkeys
                        pass
        except (OSError, WindowsError, FileNotFoundError):
            pass

        return None

    @classmethod
    def list_installed_apps(cls) -> Dict[str, str]:
        """List all registered applications in the registry."""
        apps = {}
        for hive, reg_path in cls.REG_PATHS:
            try:
                with winreg.OpenKey(hive, reg_path) as root_key:
                    i = 0
                    while True:
                        try:
                            subkey_name = winreg.EnumKey(root_key, i)
                            try:
                                with winreg.OpenKey(root_key, subkey_name) as app_key:
                                    path, _ = winreg.QueryValueEx(app_key, "")
                                    if os.path.exists(path):
                                        apps[subkey_name] = path
                            except (OSError, WindowsError):
                                pass
                            i += 1
                        except OSError:
                            break
            except (OSError, WindowsError, FileNotFoundError):
                continue
        return apps


class LnkResolver:
    """Resolve Windows shortcut (.lnk) files."""

    # Common Start Menu locations
    START_MENU_PATHS = [
        os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        os.path.join(os.environ.get('PROGRAMDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        os.path.join(os.environ.get('ALLUSERSPROFILE', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
    ]

    @classmethod
    def find_shortcut(cls, app_name: str) -> Optional[str]:
        """
        Find a shortcut by application name in Start Menu.

        Args:
            app_name: Name of the application (e.g., "Chrome", "Notepad")

        Returns:
            Path to the .lnk file if found, None otherwise.
        """
        if not PYWIN32_AVAILABLE:
            logger.warning("pywin32 is not available. Cannot search for shortcuts.")
            return None

        search_terms = [app_name.lower()]
        # Add variations
        if not app_name.endswith('.lnk'):
            search_terms.append(app_name.lower() + '.lnk')
        if not app_name.endswith('.exe'):
            search_terms.append(app_name.lower() + '.exe')

        for start_menu_path in cls.START_MENU_PATHS:
            if not os.path.exists(start_menu_path):
                continue

            for root, dirs, files in os.walk(start_menu_path):
                for file in files:
                    if file.lower().endswith('.lnk'):
                        file_lower = file.lower()
                        # Check if any search term matches
                        for term in search_terms:
                            if term in file_lower:
                                lnk_path = os.path.join(root, file)
                                return lnk_path

        return None

    @staticmethod
    def resolve_lnk(lnk_path: str) -> Optional[str]:
        """
        Resolve a .lnk file to its target executable path.

        Args:
            lnk_path: Path to the .lnk file

        Returns:
            Path to the target executable if successful, None otherwise.
        """
        if not PYWIN32_AVAILABLE:
            logger.warning("pywin32 is not available. Cannot resolve .lnk files.")
            return None

        if not os.path.exists(lnk_path):
            return None

        # Initialize COM
        pythoncom.CoInitialize()
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(lnk_path)
            target_path = shortcut.TargetPath

            # Cleanup
            del shortcut
            del shell

            return target_path if os.path.exists(target_path) else None
        except Exception as e:
            logger.error(f"Error resolving .lnk file {lnk_path}: {e}")
            return None
        finally:
            # Ensure COM is uninitialized
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @classmethod
    def resolve_app_via_shortcut(cls, app_name: str) -> Optional[str]:
        """
        Find and resolve a shortcut for an application.

        Args:
            app_name: Name of the application

        Returns:
            Path to the executable if found and resolved, None otherwise.
        """
        lnk_path = cls.find_shortcut(app_name)
        if lnk_path:
            return cls.resolve_lnk(lnk_path)
        return None


class AppLauncher:
    """Launch applications with proper GUI visibility."""

    @staticmethod
    def launch_app(exe_path: str) -> Tuple[bool, str]:
        """
        Launch an application using explorer.exe to ensure GUI visibility.

        Args:
            exe_path: Path to the executable

        Returns:
            Tuple of (success, message)
        """
        if not os.path.exists(exe_path):
            return False, f"应用程序未找到: {exe_path}"

        try:
            # Use explorer.exe to launch the application for proper GUI visibility
            # explorer.exe will handle UAC prompts and GUI context properly
            cmd = f'explorer.exe "{exe_path}"'
            result = subprocess.run(
                cmd,
                shell=True,
                capture_output=True,
                text=True,
                timeout=10
            )

            if result.returncode == 0:
                return True, f"应用程序已启动: {exe_path}"
            else:
                # Try direct execution as fallback
                try:
                    subprocess.Popen([exe_path], shell=True)
                    return True, f"应用程序已启动 (直接执行): {exe_path}"
                except Exception as e:
                    return False, f"权限被拒绝或执行失败: {str(e)}"

        except subprocess.TimeoutExpired:
            # Explorer might still be working
            return True, f"应用程序启动中: {exe_path}"
        except Exception as e:
            return False, f"启动失败: {str(e)}"


@register("opensoftware", "YourName", "通过Windows注册表和开始菜单快捷方式打开应用程序", "1.0.0")
class OpenSoftwarePlugin(Star):
    def __init__(self, context: Context):
        super().__init__(context)
        self.registry_searcher = RegistrySearcher()
        self.lnk_resolver = LnkResolver()
        self.app_launcher = AppLauncher()

    async def initialize(self):
        """插件初始化"""
        logger.info("OpenSoftware插件已初始化")

    @filter.command("open")
    async def open_app(self, event: AstrMessageEvent):
        """
        打开应用程序

        用法: /open <应用程序名称>
        示例: /open chrome
               /open notepad.exe
               /open "Microsoft Word"
        """
        message_str = event.message_str.strip()

        # Extract app name from command
        # Command format: /open appname
        parts = message_str.split(maxsplit=1)
        if len(parts) < 2:
            yield event.plain_result("用法: /open <应用程序名称>")
            return

        app_name = parts[1].strip('"\'')

        if not app_name:
            yield event.plain_result("请输入应用程序名称")
            return

        logger.info(f"尝试打开应用程序: {app_name}")

        # Step 1: Try registry search
        exe_path = RegistrySearcher.search_app(app_name)
        source = "注册表"

        # Step 2: If not found in registry, try shortcut resolution
        if not exe_path and PYWIN32_AVAILABLE:
            exe_path = LnkResolver.resolve_app_via_shortcut(app_name)
            source = "开始菜单快捷方式"

        # Step 3: If still not found, try direct path
        if not exe_path:
            # Check if it's already a full path
            if os.path.exists(app_name):
                exe_path = app_name
                source = "直接路径"
            elif os.path.exists(app_name + '.exe'):
                exe_path = app_name + '.exe'
                source = "直接路径"

        if not exe_path:
            yield event.plain_result(f"应用程序未找到: {app_name}")
            return

        # Launch the application
        success, message = AppLauncher.launch_app(exe_path)

        if success:
            yield event.plain_result(f"✓ {message}\n(来源: {source})")
        else:
            yield event.plain_result(f"✗ {message}")

    @filter.command("listapps")
    async def list_apps(self, event: AstrMessageEvent):
        """列出所有在注册表中注册的应用程序"""
        apps = RegistrySearcher.list_installed_apps()

        if not apps:
            yield event.plain_result("未找到注册的应用程序")
            return

        # Format the list
        app_list = "\n".join([f"- {name}: {path}" for name, path in list(apps.items())[:20]])  # Limit to 20

        if len(apps) > 20:
            app_list += f"\n... 以及 {len(apps) - 20} 个其他应用程序"

        yield event.plain_result(f"注册的应用程序 ({len(apps)} 个):\n{app_list}")

    async def terminate(self):
        """插件销毁"""
        logger.info("OpenSoftware插件已销毁")