
import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from core.version import get_channel
from ui.main_window import MainWindow


def _app_icon_path() -> str | None:
    """定位程序图标：打包后优先资源目录(_MEIPASS)，其次 exe 同级 assets。"""
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "assets" / "app.ico")
    candidates.append(Path(sys.executable).resolve().parent / "assets" / "app.ico")
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def main():
    # 登记启动渠道：环境变量 RELEASE_CHANNEL / 命令行参数 gitee|github，默认 github
    get_channel()
    app = QApplication(sys.argv)
    app.setApplicationName("Excel 批量数据统计工具")
    app.setOrganizationName("ExcelDataTool")
    app.setStyle("Fusion")
    icon = _app_icon_path()
    if icon:
        app.setWindowIcon(QIcon(icon))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
