# -*- coding: utf-8 -*-
"""
版本与更新配置（全项目唯一版本号源头）。

【发版说明】
以后更新脚本后重新打包，只需把下面 APP_VERSION 改成新版本号，例如 "1.0.2"，
其余位置（界面标题、右上角标签、打包 exe 文件名、更新版本比对）都会自动跟随。

配套发版约定：
1. 改完 APP_VERSION 后重新用 PyInstaller 打包；
2. 在 GitHub 仓库发布 Release，Tag 填写 v{APP_VERSION}（例如 v1.0.2）；
3. 上传的资产文件名必须符合 exe_name() 返回的规则，
   即 ExcelDataTool_v1.0.2_windows_x64.exe。
"""

# ====================== 版本号（发版只改这里） ======================
# 纯数字三段：主版本.次版本.修订号，不要带前缀 v
APP_VERSION = "1.0.1"

# ====================== GitHub 仓库信息 ======================
REPO_OWNER = "yasmin-mzsx"
REPO_NAME = "ExcelDataTool"

# 平台标识：当前只提供 Windows 64 位可执行文件
PLATFORM_TAG = "windows_x64"


def exe_name(version: str = APP_VERSION) -> str:
    """按规则拼出打包产物文件名（不含后缀 .exe）。

    例如 APP_VERSION = "1.0.2" -> "ExcelDataTool_v1.0.2_windows_x64"
    """
    return f"ExcelDataTool_v{version}_{PLATFORM_TAG}"


def release_tag(version: str = APP_VERSION) -> str:
    """GitHub Release Tag 格式，带前缀 v。"""
    return f"v{version}"
