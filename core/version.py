# -*- coding: utf-8 -*-
"""
版本与更新配置（全项目唯一版本号 / 更新渠道源头）。

【发版说明】
1. 源码调试：
   - 默认走 github；命令行可同时带渠道与版本号，顺序不限：
       python main.py                  # github + 兜底版本
       python main.py gitee            # gitee + 兜底版本
       python main.py gitee 1.1.2      # gitee + 1.1.2
       python main.py 1.1.2 gitee      # 同上，顺序不限
2. 打包发版（环境变量覆盖，打包那一刻写进 exe；不读命令行参数）：
       $env:RELEASE_CHANNEL="gitee"; $env:APP_VERSION="1.1.2"
       pyinstaller --noconfirm --clean ExcelDataTool.spec
   - 不带环境变量时：渠道默认 github，版本号取下面的 DEFAULT_VERSION。
3. 改完版本号后重新打包，并在对应平台 Release 上传同名 exe：
   文件名规则见 exe_name()，即 ExcelDataTool_v1.1.2_windows_x64.exe。
"""

import os
import re
import sys

# 命令行版本号格式：1.1 或 1.1.2（纯数字点分段，不带 v 前缀/后缀）
_VERSION_RE = re.compile(r"^\d+\.\d+(?:\.\d+)?$")

# ====================== 版本号兜底默认值（发版可改这里） ======================
# 纯数字三段：主版本.次版本.修订号，不要带前缀 v
# 打包时若设置了环境变量 APP_VERSION，以环境变量为准。
DEFAULT_VERSION = "1.1"

# ====================== 更新渠道配置表 ======================
# owner/repo：仓库归属与仓库名；api_template：最新 Release 接口模板。
CHANNELS = {
    "github": {
        "owner": "yasmin-mzsx",
        "repo": "ExcelDataTool",
        "api_template": "https://api.github.com/repos/{owner}/{repo}/releases/latest",
    },
    "gitee": {
        "owner": "zylvin",
        "repo": "ExcelDataTool",
        "api_template": "https://gitee.com/api/v5/repos/{owner}/{repo}/releases/latest",
    },
}

# 默认渠道：不传任何参数 / 环境变量时走这里
DEFAULT_CHANNEL = "github"

# 平台标识：当前只提供 Windows 64 位可执行文件
PLATFORM_TAG = "windows_x64"


def _parse_argv() -> tuple[str | None, str | None]:
    """从命令行参数里识别渠道词与版本号。

    返回 (channel, version)，识别不到的项为 None。
    渠道词：github / gitee；版本号：1.1 或 1.1.2。顺序不限，非法词忽略。
    """
    channel: str | None = None
    version: str | None = None
    for arg in sys.argv[1:]:
        candidate = arg.strip().lower()
        if candidate in CHANNELS:
            channel = candidate
        elif _VERSION_RE.match(arg.strip()):
            version = arg.strip()
    return channel, version


def _detect_channel() -> str:
    """解析当前更新渠道。

    优先级：环境变量 RELEASE_CHANNEL > 命令行渠道词 > 默认渠道。
    非法取值一律回落 DEFAULT_CHANNEL。
    """
    env_channel = os.environ.get("RELEASE_CHANNEL", "").strip().lower()
    if env_channel in CHANNELS:
        return env_channel

    argv_channel, _ = _parse_argv()
    return argv_channel or DEFAULT_CHANNEL


def get_channel() -> str:
    """返回当前生效渠道名（github / gitee）。"""
    return _detect_channel()


def get_app_version() -> str:
    """返回当前版本号。

    优先级：环境变量 APP_VERSION > 命令行版本号 > 代码兜底 DEFAULT_VERSION。
    """
    env_version = os.environ.get("APP_VERSION", "").strip()
    if env_version:
        return env_version

    _, argv_version = _parse_argv()
    return argv_version or DEFAULT_VERSION


def get_repo(channel: str | None = None) -> dict:
    """取指定渠道的仓库配置；不传则用当前生效渠道。"""
    name = channel or get_channel()
    return CHANNELS.get(name, CHANNELS[DEFAULT_CHANNEL])


def get_release_api(channel: str | None = None) -> str:
    """拼出指定渠道的最新 Release 检查接口地址。"""
    repo = get_repo(channel)
    return repo["api_template"].format(owner=repo["owner"], repo=repo["repo"])


# ====================== 对外兼容常量 / 函数 ======================
# 运行时版本号（下游界面、更新比对、打包文件名统一从这里取）
APP_VERSION = get_app_version()


def exe_name(version: str = APP_VERSION) -> str:
    """按规则拼出打包产物文件名（不含后缀 .exe）。

    例如 APP_VERSION = "1.0.2" -> "ExcelDataTool_v1.0.2_windows_x64"
    """
    return f"ExcelDataTool_v{version}_{PLATFORM_TAG}"


def release_tag(version: str = APP_VERSION) -> str:
    """Release Tag 格式，带前缀 v。"""
    return f"v{version}"
