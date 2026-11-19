# -*- coding: utf-8 -*-
# ============================================================
# 📌【更新模块 START】
# 自动/手动更新逻辑：检查 GitHub 最新 Release -> 比对版本 -> 下载新 exe
# -> 生成批处理替换旧 exe 并重启。
# 本模块只做后台线程与系统操作，不操作任何 Qt 界面控件。
# 网络请求全部使用 Python 标准库 urllib，不引入第三方依赖。
# ============================================================

import json
import os
import re
import sys
import tempfile
import urllib.request
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from core.version import (
    APP_VERSION,
    REPO_OWNER,
    REPO_NAME,
    PLATFORM_TAG,
    exe_name,
)


# GitHub 公开接口：获取仓库最新 Release
RELEASE_API = (
    f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
)

# 请求时带一个浏览器风格 User-Agent，避免被 GitHub 拦截
HTTP_HEADERS = {
    "User-Agent": f"{REPO_NAME}-updater/1.0",
    "Accept": "application/vnd.github+json",
}

# 下载超时（秒）
HTTP_TIMEOUT = 15


def _normalize_version(tag: str) -> str:
    """去掉 GitHub Tag 前缀 v/V，得到纯版本号字符串。

    例如 "v1.0.2" -> "1.0.2"；无法识别时原样返回。
    """
    return re.sub(r"^[vV]", "", tag.strip())


def _version_tuple(version: str) -> tuple:
    """把版本号字符串拆成可比较的整数元组。

    例如 "1.10.2" -> (1, 10, 2)；非数字段按 0 处理，缺段补 0。
    只比较主.次.修订三段，避免字符串比较把 "1.10.0" 误判为小于 "1.9.0"。
    """
    parts = []
    for segment in version.split(".")[:3]:
        match = re.search(r"\d+", segment)
        parts.append(int(match.group()) if match else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)


def is_newer(latest_tag: str, current: str = APP_VERSION) -> bool:
    """判断线上版本是否比当前版本新。解析失败一律按"无新版"处理。"""
    try:
        return _version_tuple(_normalize_version(latest_tag)) > _version_tuple(current)
    except Exception:
        return False


def _build_request(url: str) -> urllib.request.Request:
    return urllib.request.Request(url, headers=HTTP_HEADERS)


class CheckUpdateWorker(QThread):
    """后台线程：请求 GitHub 最新 Release，判断是否有新版。

    回传信号：
        sig_result(has_update: bool, latest_version: str,
                   notes: str, asset_url: str, message: str)
        - has_update  : 是否存在可更新的新版本
        - latest_version: 线上版本号（已去 v 前缀）
        - notes       : Release 更新日志（body，可能为空）
        - asset_url   : 匹配到的 exe 下载地址（无新版时为空串）
        - message     : 失败/提示信息，供 UI 记日志
    """

    sig_result = Signal(bool, str, str, str, str)

    def run(self):
        try:
            with urllib.request.urlopen(
                _build_request(RELEASE_API), timeout=HTTP_TIMEOUT
            ) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except Exception as exc:
            # 网络不通、仓库无 Release、接口异常等：不打断业务，回传失败信息
            self.sig_result.emit(False, APP_VERSION, "", "", f"检查更新失败：{exc}")
            return

        latest_tag = payload.get("tag_name", "")
        notes = payload.get("body", "") or ""

        if not latest_tag:
            self.sig_result.emit(False, APP_VERSION, "", "", "未获取到版本信息")
            return

        if not is_newer(latest_tag):
            self.sig_result.emit(False, _normalize_version(latest_tag), notes, "", "当前已是最新版本")
            return

        # 在 Release 资产里找匹配当前平台的 exe：文件名规则 exe_name(版本) + .exe
        want_stem = exe_name(_normalize_version(latest_tag))
        asset_url = ""
        for asset in payload.get("assets", []):
            name = asset.get("name", "")
            if name == f"{want_stem}.exe" and PLATFORM_TAG in name:
                asset_url = asset.get("browser_download_url", "")
                break

        if not asset_url:
            # 有新版但没找到对应安装包：提示业务手动获取
            self.sig_result.emit(
                True,
                _normalize_version(latest_tag),
                notes,
                "",
                f"发现新版本 {latest_tag}，但暂未上传 {PLATFORM_TAG} 更新包",
            )
            return

        self.sig_result.emit(
            True, _normalize_version(latest_tag), notes, asset_url, "发现新版本"
        )


class DownloadUpdateWorker(QThread):
    """后台线程：流式下载新 exe 到临时目录，回传进度与最终路径。

    回传信号：
        sig_progress(percent: int)        下载进度 0~100
        sig_finished(ok: bool, path: str, message: str)
            ok=True 时 path 为下载好的新 exe 临时路径；否则 message 为失败原因
    """

    sig_progress = Signal(int)
    sig_finished = Signal(bool, str, str)

    def __init__(self, asset_url: str, parent=None):
        super().__init__(parent)
        self.asset_url = asset_url

    def run(self):
        try:
            # 临时文件固定后缀 .exe，便于后续批处理识别
            fd, tmp_path = tempfile.mkstemp(
                prefix="ExcelDataTool_update_", suffix=".exe"
            )
            os.close(fd)

            with urllib.request.urlopen(
                _build_request(self.asset_url), timeout=HTTP_TIMEOUT
            ) as resp:
                total = int(resp.headers.get("Content-Length", 0))
                downloaded = 0
                with open(tmp_path, "wb") as f:
                    while True:
                        chunk = resp.read(1024 * 256)  # 256KB 一块
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total > 0:
                            self.sig_progress.emit(min(100, int(downloaded * 100 / total)))
                        else:
                            # 服务器没给总大小时，只回传不确定进度，避免卡死在 0
                            self.sig_progress.emit(50)

            # 【完整性校验】下载结束后核对文件大小，残缺/空包不进入替换环节
            local_size = os.path.getsize(tmp_path)
            if local_size <= 0:
                os.remove(tmp_path)
                self.sig_finished.emit(False, "", "下载更新包为空文件，已中止更新")
                return
            if total > 0 and local_size != total:
                os.remove(tmp_path)
                self.sig_finished.emit(
                    False, "", f"更新包不完整（已下载 {local_size} / 共 {total} 字节），已中止更新"
                )
                return

            self.sig_progress.emit(100)
            self.sig_finished.emit(True, tmp_path, "下载完成")
        except Exception as exc:
            self.sig_finished.emit(False, "", f"下载更新包失败：{exc}")


def prepare_update_script(new_exe_path: str, old_pid: int | None = None) -> str:
    """生成安全的 onefile 更新批处理脚本。"""
    old_exe = Path(sys.executable).resolve()
    new_exe = Path(new_exe_path).resolve()
    backup_exe = old_exe.with_suffix(old_exe.suffix + ".bak")
    bat_path = Path(tempfile.gettempdir()) / "ExcelDataTool_apply_update.bat"
    pid = int(old_pid or os.getpid())

    bat_content = f'''@echo off
chcp 65001 >nul
setlocal EnableExtensions
set "OLD_EXE={old_exe}"
set "NEW_EXE={new_exe}"
set "BACKUP_EXE={backup_exe}"
set "OLD_PID={pid}"
set "APP_NAME={old_exe.stem}"

:: Wait for the exact old process to exit. Never force replacement on timeout.
set /a WAIT_COUNT=0
:WAIT_OLD_PROCESS
powershell -NoProfile -Command "if (Get-Process -Id %OLD_PID% -ErrorAction SilentlyContinue) {{ exit 1 }} else {{ exit 0 }}" >nul 2>nul
if not errorlevel 1 goto OLD_PROCESS_EXITED
set /a WAIT_COUNT+=1
if %WAIT_COUNT% GEQ 120 goto WAIT_FAILED
timeout /t 1 /nobreak >nul
goto WAIT_OLD_PROCESS

:OLD_PROCESS_EXITED
if exist "%BACKUP_EXE%" del /F /Q "%BACKUP_EXE%" >nul 2>nul
copy /Y "%OLD_EXE%" "%BACKUP_EXE%" >nul
if errorlevel 1 goto REPLACE_FAILED
copy /Y "%NEW_EXE%" "%OLD_EXE%" >nul
if errorlevel 1 goto ROLLBACK

for %%A in ("%NEW_EXE%") do set "NEW_SIZE=%%~zA"
for %%B in ("%OLD_EXE%") do set "OLD_SIZE=%%~zB"
if not "%NEW_SIZE%"=="%OLD_SIZE%" goto ROLLBACK
if "%OLD_SIZE%"=="0" goto ROLLBACK

:: Start the new build as a fresh PyInstaller instance.
set "PYINSTALLER_RESET_ENVIRONMENT=1"
start "" "%OLD_EXE%"

:: Keep the backup until the new process is detected.
set /a START_COUNT=0
:WAIT_NEW_PROCESS
powershell -NoProfile -Command "if (Get-Process -Name '%APP_NAME%' -ErrorAction SilentlyContinue) {{ exit 0 }} else {{ exit 1 }}" >nul 2>nul
if not errorlevel 1 goto UPDATE_SUCCESS
set /a START_COUNT+=1
if %START_COUNT% GEQ 15 goto ROLLBACK
timeout /t 1 /nobreak >nul
goto WAIT_NEW_PROCESS

:UPDATE_SUCCESS
del /F /Q "%BACKUP_EXE%" >nul 2>nul
goto CLEANUP

:ROLLBACK
if exist "%BACKUP_EXE%" (
    copy /Y "%BACKUP_EXE%" "%OLD_EXE%" >nul
    if not errorlevel 1 del /F /Q "%BACKUP_EXE%" >nul 2>nul
)
goto CLEANUP

:REPLACE_FAILED
if exist "%BACKUP_EXE%" del /F /Q "%BACKUP_EXE%" >nul 2>nul
goto CLEANUP

:WAIT_FAILED
echo [%date% %time%] Update cancelled: old PID %OLD_PID% did not exit.>>"%TEMP%\\ExcelDataTool_update.log"
goto CLEANUP

:CLEANUP
if exist "%NEW_EXE%" del /F /Q "%NEW_EXE%" >nul 2>nul
if exist "%~f0" del /F /Q "%~f0" >nul 2>nul
endlocal
'''
    bat_path.write_text(bat_content, encoding="utf-8")
    return str(bat_path)

def launch_update_and_exit(new_exe_path: str):
    """生成批处理并退出主程序，由批处理完成替换与重启。"""
    old_pid = os.getpid()
    bat_path = prepare_update_script(new_exe_path, old_pid=old_pid)
    # detached 方式启动批处理，主程序随后退出
    os.startfile(bat_path)
    sys.exit(0)


# ============================================================
# 📌【更新模块 END】
# ============================================================
