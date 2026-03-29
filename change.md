# 变更记录 (Change Log)

## 版本 1.0.1 - 2026-03-29

### 概述
本次更新主要针对代码安全性、性能和代码质量进行了全面优化，解决了代码评审团指出的关键问题。

### 🛡️ 安全修复

#### 1. 命令注入安全漏洞修复
**问题**: `AppLauncher.launch_app()` 方法使用 `shell=True` 和字符串拼接执行命令，存在命令注入风险。

**修复**:
- 改用 `asyncio.create_subprocess_exec(["explorer.exe", exe_path])` 避免命令拼接
- 直接执行时使用 `asyncio.create_subprocess_exec(exe_path)` 无 shell 调用
- 方法改为异步 `async def launch_app()`，避免阻塞机器人事件循环

**代码评审团反馈**:
> Linus Torvalds: "看看这段代码里的 subprocess.run(..., shell=True)！你是在写现代异步机器人，还是在上世纪90年代写漏洞百出的批处理脚本？不要用 shell=True，用 asyncio.create_subprocess_exec！重写它！"

#### 2. 注册表值类型安全校验
**问题**: `winreg.QueryValueEx()` 可能返回非字符串类型，直接使用会导致 `TypeError`。

**修复**:
- 只接受 `REG_SZ` 和 `REG_EXPAND_SZ` 字符串类型
- 对 `REG_EXPAND_SZ` 使用 `os.path.expandvars()` 展开环境变量
- 非字符串类型自动跳过，避免程序崩溃

### ⚡ 性能与功能改进

#### 1. 缓存状态同步机制
**问题**: `RegistrySearcher.list_installed_apps()` 使用 `@lru_cache(maxsize=1)` 永久缓存，运行时安装新软件无法感知。

**修复**:
- 添加 `RegistrySearcher.clear_cache()` 类方法
- 新增 `/refreshapps` 命令手动刷新缓存
- 用户可在安装新软件后执行该命令更新应用列表

**代码评审团反馈**:
> GLaDOS: "我注意到你给读取已安装应用的函数加上了一个永不过期的 @lru_cache。我猜你的逻辑是：‘人类总是如此懒惰，他们绝对不可能在机器人运行期间安装新软件。’ 真是令人赞叹的乐观主义。"

#### 2. 异步兼容性优化
**问题**: 同步 `subprocess.run(..., timeout=10)` 会阻塞异步事件循环。

**修复**:
- 改用 `asyncio.create_subprocess_exec()` 异步执行
- 使用 `asyncio.wait_for(process.wait(), timeout=10)` 处理超时
- 保持机器人在启动程序期间的响应能力

### 📝 代码质量提升

#### 1. 导入规范优化
- 将 `SimilarityMatcher.calculate_similarity()` 中的局部导入 `import difflib` 移到文件顶部
- 符合 PEP 8 导入规范

#### 2. 异常处理简化
- 所有 `except (OSError, FileNotFoundError)` 简化为 `except OSError`
- `FileNotFoundError` 已经是 `OSError` 的子类，简化代码逻辑

#### 3. 循环结构修复
- 修复了 `continue` 语句在非循环上下文中的语法错误
- 改进了注册表枚举的异常处理逻辑

### 🆕 新增功能

#### 1. `/refreshapps` 命令
- **功能**: 手动刷新已安装应用程序缓存
- **用法**: `/refreshapps`
- **场景**: 在安装新软件后使用，确保插件能识别最新安装的应用

### 🔧 技术细节

#### 修改的文件
- `main.py`: 所有安全性和性能优化

#### 主要变更点
1. **AppLauncher 类** (第405-453行):
   - `launch_app()` 方法改为异步
   - 移除 `shell=True`，使用安全的子进程调用
   - 支持异步超时处理

2. **RegistrySearcher 类** (第233-262行):
   - 新增 `clear_cache()` 静态方法
   - 注册表值类型安全校验
   - 环境变量展开支持

3. **OpenSoftwarePlugin 类** (第552-555行):
   - 新增 `refresh_apps()` 方法处理 `/refreshapps` 命令

### 📋 测试建议

1. **安全测试**:
   - 尝试使用包含特殊字符的应用路径
   - 验证命令注入防护是否生效

2. **功能测试**:
   - 安装新软件后使用 `/refreshapps` 刷新缓存
   - 验证 `/open` 命令能否找到新安装的应用

3. **性能测试**:
   - 在机器人运行期间打开大型应用，观察是否影响其他功能响应
   - 测试多个并发打开请求的处理能力

### ⚠️ 已知限制

- **平台依赖**: 插件仍然依赖 Windows 特定 API (`winreg`, `WScript.Shell`)
- **pywin32 依赖**: `.lnk` 文件解析需要 pywin32 库

**代码评审团反馈**:
> Richard Stallman: "我必须打断一下。你这里大量使用了 winreg 和 WScript.Shell，这表明这个插件被死死地绑在了一个专有的、剥夺用户自由的非自由操作系统上！"

### 📈 兼容性

- **向后兼容**: 所有现有命令和功能保持不变
- **依赖更新**: 无需新增依赖，使用现有 asyncio 库
- **Python 版本**: 支持 Python 3.7+

### 🔄 升级说明

1. 替换 `main.py` 文件
2. 无需重启机器人，插件会自动加载新代码
3. 首次使用建议执行 `/refreshapps` 刷新应用列表

---

*本次优化由代码评审团监督完成，特别感谢 Linus Torvalds、GLaDOS 和 Richard Stallman 的严格审查。*