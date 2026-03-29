# astrbot-plugin-opensoftware

AstrBot 插件：通过Windows注册表和开始菜单快捷方式打开应用程序

## 功能

- 通过 `/open <应用程序名称>` 命令打开应用程序
- 支持从Windows注册表搜索应用程序路径
- 支持解析开始菜单快捷方式 (.lnk 文件)
- 使用 `explorer.exe` 启动应用程序以确保GUI可见性
- 提供 `/listapps` 命令列出所有注册的应用程序

## 安装

### 1. 安装插件
将插件文件夹 `astrbot_plugin_OpenSoftware` 复制到 `AstrBot/data/plugins/` 目录。

### 2. 安装依赖
在插件目录中运行以下命令安装所需依赖：

```bash
cd AstrBot/data/plugins/astrbot_plugin_OpenSoftware
pip install -r requirements.txt
```

或者使用AstrBot的Python环境：
```bash
# 在AstrBot安装目录下运行
python -m pip install -r data/plugins/astrbot_plugin_OpenSoftware/requirements.txt
```

### 3. 重载插件
在AstrBot管理界面中重载插件，或使用重载命令。

## 依赖

- pywin32 (用于解析.lnk文件)

## 使用方法

- `/open chrome` - 打开Chrome浏览器
- `/open notepad` - 打开记事本
- `/open "Microsoft Word"` - 打开Microsoft Word
- `/listapps` - 列出所有在注册表中注册的应用程序

## 工作原理

1. 首先在Windows注册表中搜索应用程序路径 (HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths 等)
2. 如果在注册表中未找到，则搜索开始菜单快捷方式
3. 使用 `explorer.exe` 启动应用程序以确保正确的GUI上下文和UAC处理

## 错误处理

- 应用程序未找到时返回友好错误信息
- 权限被拒绝时提供相应提示
- 优雅处理依赖缺失情况

## 支持

- [AstrBot 仓库](https://github.com/AstrBotDevs/AstrBot)
- [AstrBot 插件开发文档 (中文)](https://docs.astrbot.app/dev/star/plugin-new.html)
- [AstrBot 插件开发文档 (英文)](https://docs.astrbot.app/en/dev/star/plugin-new.html)
