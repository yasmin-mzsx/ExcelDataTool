# -*- mode: python ; coding: utf-8 -*-

# 【更新模块】从版本模块读取版本号，发版时只需改 core/version.py
import sys
sys.path.insert(0, '.')
from core.version import APP_VERSION

# 打包产物名：ExcelDataTool_v1.0.1_windows_x64.exe（版本号自动跟随 APP_VERSION）
EXE_NAME = f'ExcelDataTool_v{APP_VERSION}_windows_x64'

excludes = [
    # WebEngine 浏览器内核 / WebChannel / Pdf
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineQuick',
    'PySide6.QtWebChannel', 'PySide6.QtPdf', 'PySide6.QtPdfWidgets', 'PySide6.QtPdfQuick',
    # QML / Quick / Quick3D / 控件样式
    'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtQuick3D', 'PySide6.QtQuickControls2',
    'PySide6.QtQuickWidgets', 'PySide6.QtQuickTest', 'PySide6.QtQuickLayouts',
    'PySide6.QtQuickDialogs2', 'PySide6.QtQuickShapes', 'PySide6.QtQuickTemplates2',
    'PySide6.QtQuickEffects', 'PySide6.QtQuickParticles', 'PySide6.QtQuickTimeline',
    'PySide6.QtQuickVectorImage', 'PySide6.QtQuickNativeStyle',
    # 多媒体 + FFmpeg
    'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets', 'PySide6.QtSpatialAudio',
    # Qt 3D
    'PySide6.Qt3DCore', 'PySide6.Qt3DAnimation', 'PySide6.Qt3DExtras',
    'PySide6.Qt3DInput', 'PySide6.Qt3DLogic', 'PySide6.Qt3DRender',
    # 图表 / 数据可视化
    'PySide6.QtCharts', 'PySide6.QtDataVisualization', 'PySide6.QtGraphs',
    # 其余未使用的 Qt 功能模块
    'PySide6.QtDesigner', 'PySide6.QtHelp', 'PySide6.QtLocation', 'PySide6.QtPositioning',
    'PySide6.QtBluetooth', 'PySide6.QtNfc', 'PySide6.QtSensors', 'PySide6.QtSerialBus',
    'PySide6.QtSerialPort', 'PySide6.QtRemoteObjects', 'PySide6.QtScxml', 'PySide6.QtSql',
    'PySide6.QtTest', 'PySide6.QtTextToSpeech', 'PySide6.QtVirtualKeyboard',
    'PySide6.QtWebSockets', 'PySide6.QtWebView', 'PySide6.QtLottie', 'PySide6.QtStateMachine',
    'PySide6.QtUiTools', 'PySide6.QtAxContainer', 'PySide6.QtNetworkAuth',
    'PySide6.QtHttpServer', 'PySide6.QtDBus', 'PySide6.QtPrintSupport',
    'PySide6.QtConcurrent', 'PySide6.QtSvg', 'PySide6.QtSvgWidgets',
    # 无关第三方（本项目未使用，避免被间接带入）
    'xlrd', 'matplotlib', 'PIL', 'pytest', 'IPython', 'jupyter', 'notebook',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/app.ico', 'assets'), ('assets/icon_upgrade.svg', 'assets')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets/app.ico'],
)
