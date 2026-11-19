# -*- coding: utf-8 -*-
"""
版本与更新配置（全项目唯一版本号 / 更新渠道源头）。

配置文件：config/app_channel.json，字段：
  channel : "github" | "gitee"   更新渠道
  version : "1.1" | "1.1.2"       版本号（纯数字点分段，不带 v 前缀）

无论是源码运行还是打包后的 exe，都以该 JSON 为准；
JSON 缺失或字段非法时，回落兜底：channel=github，version=DEFAULT_VERSION。
"""

import json
import os
import sys
from pathlib import Path

# ====================== 版本号兜底默认值（JSON 缺失时使用） ======================
DEFAULT_VERSION = "1.1"

# 默认渠道：JSON 缺失或非法时使用
DEFAULT_CHANNEL = "github"

# 平台标识：当前只提供 Windows 64 位可执行文件
PLATFORM_TAG = "windows_x64"

# ====================== 更新渠道配置表 ======================
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

# 配置文件名（位于 config/ 目录下）
_CONFIG_DIR = "config"
_CONFIG_FILE = "app_channel.json"


def _config_candidates() -> list[Path]:
    """按优先级返回配置文件候选路径。

    1. 打包后 exe 同级目录的 config/（外部可改，便携覆盖）；
    2. PyInstaller 解压目录 _MEIPASS 内打包副本；
    3. 源码运行时项目根目录下的 config/。
    """
    candidates: list[Path] = []

    # 1. exe 同级目录（打包后绿色便携）
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / _CONFIG_DIR / _CONFIG_FILE)

    # 2. PyInstaller 解压目录内副本
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / _CONFIG_DIR / _CONFIG_FILE)

    # 3. 源码项目根目录
    candidates.append(Path(__file__).resolve().parents[1] / _CONFIG_DIR / _CONFIG_FILE)

    return candidates


def _load_config() -> dict:
    """读取渠道配置 JSON；读不到或解析失败时返回空字典。"""
    for path in _config_candidates():
        try:
            if path.exists():
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            continue
    return {}


def get_channel() -> str:
    """返回当前生效渠道名（github / gitee），非法值回落 DEFAULT_CHANNEL。"""
    channel = str(_load_config().get("channel", "")).strip().lower()
    return channel if channel in CHANNELS else DEFAULT_CHANNEL


def get_app_version() -> str:
    """返回当前版本号；JSON 缺失或非法时回落 DEFAULT_VERSION。"""
    version = str(_load_config().get("version", "")).strip()
    return version or DEFAULT_VERSION


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
