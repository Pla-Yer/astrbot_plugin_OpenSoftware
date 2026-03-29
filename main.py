import os
import subprocess
import warnings
import winreg
from functools import lru_cache

# Try to import pywin32 for .lnk parsing
try:
    import pythoncom
    import win32com.client

    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False
    warnings.warn("pywin32 is not installed. LnkResolver will not work.")

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.star import Context, Star, register


class SimilarityMatcher:
    """字符串相似度匹配工具类"""

    # 相似度阈值
    EXACT_THRESHOLD = 1.0
    HIGH_SIMILARITY_THRESHOLD = 0.9
    MEDIUM_SIMILARITY_THRESHOLD = (
        0.86  # 提高阈值以避免todesk匹配autodesk (0.857 < 0.86)
    )

    @staticmethod
    @lru_cache(maxsize=1024)
    def calculate_similarity(term: str, candidate: str) -> float:
        """计算两个字符串的相似度分数（0-1），带缓存"""
        import difflib

        return difflib.SequenceMatcher(None, term.lower(), candidate.lower()).ratio()

    @classmethod
    def find_best_match(
        cls, term: str, candidates: list[str], include_paths: bool = False
    ) -> tuple[str, float, str | None] | None:
        """
        在候选列表中查找最佳匹配

        Args:
            term: 搜索词
            candidates: 候选字符串列表
            include_paths: 是否包含路径信息（候选为(name, path)元组）

        Returns:
            (最佳匹配名称, 相似度分数, 路径或None)
        """
        if not candidates:
            return None

        best_match = None
        best_score = 0.0
        best_path = None

        for candidate in candidates:
            if include_paths:
                cand_name, cand_path = candidate
            else:
                cand_name = candidate
                cand_path = None

            # 检查精确匹配（最高优先级）
            if cand_name.lower() == term.lower():
                return (cand_name, 1.0, cand_path)

            # 检查前缀匹配（次高优先级）
            if cand_name.lower().startswith(term.lower()):
                # 前缀匹配视为高相似度匹配，分数设为0.95
                return (cand_name, 0.95, cand_path)

            score = cls.calculate_similarity(term, cand_name)

            # 优先级：精确匹配 > 高相似度 > 中等相似度
            if score == cls.EXACT_THRESHOLD:
                return (cand_name, score, cand_path)
            elif score > best_score:
                best_match = cand_name
                best_score = score
                best_path = cand_path

        # 检查是否达到最低相似度阈值
        if best_score >= cls.MEDIUM_SIMILARITY_THRESHOLD:
            return (best_match, best_score, best_path)

        return None


class RegistrySearcher:
    """Search for installed applications in Windows Registry."""

    # Common registry paths where applications are registered
    REG_PATHS = [
        (
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
        ),
        (
            winreg.HKEY_CURRENT_USER,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
        ),
        (
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths",
        ),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\Applications"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Classes\Applications"),
    ]

    @classmethod
    def search_app(cls, app_name: str) -> str | None:
        """
        Search for an application by name in Windows Registry.

        Args:
            app_name: Name of the application (e.g., "chrome.exe", "notepad.exe")

        Returns:
            Full path to the executable if found, None otherwise.
        """
        # 首先尝试精确匹配（保持向后兼容）
        exact_path = cls._search_exact(app_name)
        if exact_path:
            return exact_path

        # 如果精确匹配失败，尝试相似度匹配
        return cls._search_by_similarity(app_name)

    @classmethod
    def _search_exact(cls, app_name: str) -> str | None:
        """精确匹配搜索（原search_app的逻辑）"""
        # Try with .exe extension if not already present
        if not app_name.lower().endswith(".exe"):
            app_name_exe = app_name + ".exe"
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

            except (OSError, FileNotFoundError):
                continue

        return None

    @classmethod
    def _search_by_similarity(cls, app_name: str) -> str | None:
        """通过相似度匹配搜索应用"""
        # 获取所有已安装应用
        apps = cls.list_installed_apps()
        if not apps:
            return None

        # 准备候选列表
        candidates = [(name, path) for name, path in apps.items()]

        # 查找最佳匹配
        result = SimilarityMatcher.find_best_match(
            app_name, candidates, include_paths=True
        )

        if result:
            match_name, score, match_path = result
            logger.debug(
                f"相似度匹配成功: '{app_name}' -> '{match_name}' (相似度: {score:.2f})"
            )
            return match_path

        return None

    @staticmethod
    def _try_registry_path(hive, reg_path, key_name: str) -> str | None:
        """Try to read a registry path and return the executable path."""
        try:
            with winreg.OpenKey(hive, reg_path) as root_key:
                try:
                    with winreg.OpenKey(root_key, key_name) as app_key:
                        # Read the default value (empty string) which usually contains the path
                        path, _ = winreg.QueryValueEx(app_key, "")
                        return path
                except (OSError, FileNotFoundError):
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
        except (OSError, FileNotFoundError):
            pass

        return None

    @classmethod
    @lru_cache(maxsize=1)
    def list_installed_apps(cls) -> dict[str, str]:
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
                            except OSError:
                                pass
                            i += 1
                        except OSError:
                            break
            except (OSError, FileNotFoundError):
                continue
        return apps


class LnkResolver:
    """Resolve Windows shortcut (.lnk) files."""

    # Common Start Menu locations
    START_MENU_PATHS = [
        os.path.join(
            os.environ.get("APPDATA", ""),
            "Microsoft",
            "Windows",
            "Start Menu",
            "Programs",
        ),
        os.path.join(
            os.environ.get("PROGRAMDATA", ""),
            "Microsoft",
            "Windows",
            "Start Menu",
            "Programs",
        ),
        os.path.join(
            os.environ.get("ALLUSERSPROFILE", ""),
            "Microsoft",
            "Windows",
            "Start Menu",
            "Programs",
        ),
    ]

    @classmethod
    def find_shortcut(cls, app_name: str) -> str | None:
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

        # 收集所有快捷方式名称
        shortcuts = cls._collect_shortcuts()
        if not shortcuts:
            return None

        # 查找最佳匹配
        result = SimilarityMatcher.find_best_match(
            app_name, shortcuts, include_paths=False
        )

        if result:
            match_name, score, _ = result
            # 根据匹配名称查找对应的.lnk文件路径
            for lnk_path, lnk_name in cls._get_shortcut_mapping():
                if lnk_name.lower() == match_name.lower():
                    logger.debug(
                        f"快捷方式相似度匹配成功: '{app_name}' -> '{match_name}' (相似度: {score:.2f})"
                    )
                    return lnk_path

        return None

    @classmethod
    def _collect_shortcuts(cls) -> list[str]:
        """收集所有快捷方式名称"""
        shortcuts = []

        for start_menu_path in cls.START_MENU_PATHS:
            if not os.path.exists(start_menu_path):
                continue

            for root, dirs, files in os.walk(start_menu_path):
                for file in files:
                    if file.lower().endswith(".lnk"):
                        # 移除.lnk扩展名作为候选名称
                        name = file[:-4] if file.lower().endswith(".lnk") else file
                        shortcuts.append(name)

        return shortcuts

    @classmethod
    def _get_shortcut_mapping(cls) -> list[tuple[str, str]]:
        """获取快捷方式路径和名称的映射"""
        mapping = []

        for start_menu_path in cls.START_MENU_PATHS:
            if not os.path.exists(start_menu_path):
                continue

            for root, dirs, files in os.walk(start_menu_path):
                for file in files:
                    if file.lower().endswith(".lnk"):
                        lnk_path = os.path.join(root, file)
                        name = file[:-4] if file.lower().endswith(".lnk") else file
                        mapping.append((lnk_path, name))

        return mapping

    @staticmethod
    def resolve_lnk(lnk_path: str) -> str | None:
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
    def resolve_app_via_shortcut(cls, app_name: str) -> str | None:
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
    def launch_app(exe_path: str) -> tuple[bool, str]:
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
                cmd, shell=True, capture_output=True, text=True, timeout=10
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


@register(
    "opensoftware",
    "YourName",
    "通过Windows注册表和开始菜单快捷方式打开应用程序",
    "1.0.0",
)
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

        app_name = parts[1].strip("\"'")

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
            elif os.path.exists(app_name + ".exe"):
                exe_path = app_name + ".exe"
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
        app_list = "\n".join(
            [f"- {name}: {path}" for name, path in list(apps.items())[:20]]
        )  # Limit to 20

        if len(apps) > 20:
            app_list += f"\n... 以及 {len(apps) - 20} 个其他应用程序"

        yield event.plain_result(f"注册的应用程序 ({len(apps)} 个):\n{app_list}")

    async def terminate(self):
        """插件销毁"""
        logger.info("OpenSoftware插件已销毁")
