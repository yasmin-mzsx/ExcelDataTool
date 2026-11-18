
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QGroupBox, QPlainTextEdit,
    QFileDialog, QMessageBox, QSplitter, QTableWidget, QTableWidgetItem,
    QHeaderView, QProgressBar, QCheckBox, QSizePolicy, QScrollArea, QFrame,
    QCompleter
)

# ====================== BACKEND IMPORTS START ======================
from core.process_worker import ExcelProcessWorker
# ====================== BACKEND IMPORTS END ======================


APP_NAME = "凭证附件批量处理工具"
CONFIG_DIR_NAME = "config"
CONFIG_FILE_NAME = "app_config.json"
DEFAULT_MAPPING_FILE_NAME = "default_mapping.json"


def app_base_dir() -> Path:
    """开发运行时使用项目目录；打包后使用 EXE 所在目录，便于绿色便携。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def extract_doc_key(filename: str) -> str:
    """从文件名提取单据名（固定映射键名）。

    规则：只保留中文汉字，去掉日期里的"月"字；提取结果为空时回退用文件名去后缀。
    例如：
        "01销售出库单_2-189.xlsx"   -> "销售出库单"
        "成本调整单2月172#.xlsx"    -> "成本调整单"
        "其他出库单（全部）175.xlsx" -> "其他出库单全部"
    """
    stem = filename.rsplit(".", 1)[0]
    chinese = "".join(re.findall(r"[\u4e00-\u9fff]", stem))
    chinese = chinese.replace("月", "")
    return chinese if chinese else stem


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v1.0.0")
        self.resize(1000, 600)
        self.setMinimumSize(850, 550)

        self.base_dir = app_base_dir()
        self.config_dir = self.base_dir / CONFIG_DIR_NAME
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.config_path = self.config_dir / CONFIG_FILE_NAME
        self.default_mapping_path = self.config_dir / DEFAULT_MAPPING_FILE_NAME

        self.user_config = self.load_json(self.config_path, {})
        self.default_mapping = self.load_json(self.default_mapping_path, {})
        self.field_mapping = dict(self.default_mapping)
        self.field_mapping.update(self.user_config.get("field_mapping", {}))

        # ====================== BACKEND MEMBER START ======================
        self.process_worker: ExcelProcessWorker | None = None
        # ====================== BACKEND MEMBER END ======================

        self._build_ui()
        self._load_saved_paths()
        self._refresh_file_list()
        self._refresh_mapping_table()
        self._update_start_button()

    def load_json(self, path: Path, fallback):
        try:
            if path.exists():
                with path.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                    return data if isinstance(data, dict) else fallback
        except Exception:
            pass
        return fallback

    def save_config(self):
        data = {
            "last_source_dir": self.source_edit.text().strip(),
            "last_template": self.template_edit.text().strip(),
            "field_mapping": self.field_mapping,
        }
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            with self.config_path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as exc:
            self.append_log(f"[错误] 配置保存失败：{exc}", "error")

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 14, 18, 16)
        root.setSpacing(12)

        # 顶部标题
        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel(APP_NAME)
        title.setObjectName("appTitle")
        title_box.addWidget(title)
        header.addLayout(title_box)
        header.addStretch()
        self.version_label = QLabel("V1.0.0")
        self.version_label.setObjectName("versionLabel")
        header.addWidget(self.version_label, alignment=Qt.AlignVCenter)
        root.addLayout(header)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        # 左侧操作配置区（包进滚动区：窗口压缩时内容不截断，出垂直滚动条）
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(10)

        source_group = QGroupBox("1. 选择源文件目录")
        source_layout = QGridLayout(source_group)
        self.source_edit = QLineEdit()
        self.source_edit.setPlaceholderText("请选择包含 .xlsx 文件的文件夹")
        self.source_browse_btn = QPushButton("浏览目录…")
        self.file_count_label = QLabel("尚未选择目录")
        source_layout.addWidget(self.source_edit, 0, 0)
        source_layout.addWidget(self.source_browse_btn, 0, 1)
        source_layout.addWidget(self.file_count_label, 1, 0, 1, 2)
        left_layout.addWidget(source_group)

        template_group = QGroupBox("2. 选择输出模板")
        template_layout = QGridLayout(template_group)
        self.template_edit = QLineEdit()
        self.template_edit.setPlaceholderText("请选择 .xlsx 输出模板")
        self.template_browse_btn = QPushButton("选择文件…")
        template_layout.addWidget(self.template_edit, 0, 0)
        template_layout.addWidget(self.template_browse_btn, 0, 1)
        left_layout.addWidget(template_group)

        field_group = QGroupBox("3. 字段配置")
        field_layout = QVBoxLayout(field_group)
        field_layout.setContentsMargins(12, 14, 12, 12)
        field_layout.setSpacing(8)

        selector_row = QHBoxLayout()
        selector_row.addWidget(QLabel("选择文件："))
        self.file_combo = QComboBox()
        self.file_combo.setMinimumWidth(180)
        self.file_combo.setEditable(True)
        self.file_combo.setInsertPolicy(QComboBox.NoInsert)
        self.file_combo.lineEdit().setPlaceholderText("输入关键字搜索文件…")
        self.file_combo.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.file_combo.completer().setFilterMode(Qt.MatchContains)
        selector_row.addWidget(self.file_combo, 1)
        field_layout.addLayout(selector_row)

        field_layout.addWidget(QLabel("统计字段（使用 | 分隔）："))
        self.field_edit = QLineEdit()
        self.field_edit.setPlaceholderText("例如：供应商|物料名称|单位|实退数量")
        field_layout.addWidget(self.field_edit)

        field_actions = QHBoxLayout()
        self.save_mapping_btn = QPushButton("保存当前配置")
        self.reset_mapping_btn = QPushButton("恢复默认字段")
        self.save_mapping_btn.setMinimumHeight(30)
        self.reset_mapping_btn.setMinimumHeight(30)
        field_actions.addWidget(self.save_mapping_btn)
        field_actions.addWidget(self.reset_mapping_btn)
        field_actions.addStretch()
        field_layout.addLayout(field_actions)

        stats_row = QHBoxLayout()
        self.stats_label = QLabel("全部文件配置：共 0 个文件，已匹配 0 个，未匹配 0 个")
        self.stats_label.setObjectName("statsLabel")
        stats_row.addWidget(self.stats_label)
        stats_row.addStretch()
        self.locate_btn = QPushButton("定位未匹配 ▼")
        self.locate_btn.setMinimumHeight(28)
        stats_row.addWidget(self.locate_btn)
        field_layout.addLayout(stats_row)

        self.mapping_table = QTableWidget(0, 5)
        self.mapping_table.setHorizontalHeaderLabels(["选", "#", "文件名", "单据名", "统计字段"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.mapping_table.setColumnWidth(0, 36)
        self.mapping_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.mapping_table.setColumnWidth(1, 40)
        self.mapping_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.mapping_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.mapping_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.mapping_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.mapping_table.setSelectionMode(QTableWidget.SingleSelection)
        self.mapping_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.mapping_table.setMinimumHeight(160)
        self.mapping_table.setMaximumHeight(240)
        self.mapping_table.verticalHeader().setVisible(False)
        field_layout.addWidget(self.mapping_table)

        table_actions = QHBoxLayout()
        self.select_all_btn = QPushButton("全选")
        self.delete_mapping_btn = QPushButton("删除选中配置")
        self.select_all_btn.setMinimumHeight(30)
        self.delete_mapping_btn.setMinimumHeight(30)
        self._all_checked = False  # 全选按钮toggle状态，默认全不选
        table_actions.addWidget(self.select_all_btn)
        table_actions.addWidget(self.delete_mapping_btn)
        table_actions.addStretch()
        field_layout.addLayout(table_actions)
        left_layout.addWidget(field_group)

        bottom_left = QHBoxLayout()
        self.auto_open_check = QCheckBox("处理完成后自动打开输出目录")
        self.auto_open_check.setChecked(bool(self.user_config.get("auto_open_output", False)))
        self.start_btn = QPushButton("▶  开始处理")
        self.start_btn.setObjectName("primaryButton")
        self.start_btn.setEnabled(False)
        self.start_btn.setMinimumHeight(36)
        bottom_left.addWidget(self.auto_open_check)
        bottom_left.addStretch()
        bottom_left.addWidget(self.start_btn)
        left_layout.addLayout(bottom_left)
        left_layout.addStretch()

        left_scroll.setWidget(left)

        # 右侧日志与状态区
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(10)

        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        log_top = QHBoxLayout()
        log_top.addWidget(QLabel("实时显示程序运行信息"))
        log_top.addStretch()
        self.clear_log_btn = QPushButton("清空日志")
        log_top.addWidget(self.clear_log_btn)
        log_layout.addLayout(log_top)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlaceholderText("程序运行信息将在这里显示……")
        log_layout.addWidget(self.log_view, 1)
        right_layout.addWidget(log_group, 3)

        status_group = QGroupBox("处理状态")
        status_layout = QVBoxLayout(status_group)
        self.status_label = QLabel("等待配置完成")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.progress_bar)
        stats = QHBoxLayout()
        self.success_label = QLabel("成功：0")
        self.failed_label = QLabel("失败：0")
        self.success_label.setStyleSheet("color: #1E8449; font-weight: 600;")
        self.failed_label.setStyleSheet("color: #C0392B; font-weight: 600;")
        stats.addWidget(self.success_label)
        stats.addStretch()
        stats.addWidget(self.failed_label)
        status_layout.addLayout(stats)
        right_layout.addWidget(status_group, 1)

        splitter.addWidget(left_scroll)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([700, 430])

        self._apply_style()
        self._connect_signals()

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #F5F7FB; color: #243247; font-family: "Microsoft YaHei UI"; font-size: 10pt; }
            QLabel#appTitle { font-size: 19pt; font-weight: 700; color: #172B4D; }
            QLabel#subTitle, QLabel#versionLabel { color: #718096; }
            QLabel#statsLabel { color: #52657F; font-size: 9.5pt; }
            QGroupBox { background: #FFFFFF; border: 1px solid #DCE4F0; border-radius: 9px; margin-top: 10px; padding: 12px 10px 10px 10px; font-weight: 600; }
            QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; color: #315A91; }
            QLineEdit, QComboBox, QPlainTextEdit, QTableWidget { background: #FFFFFF; border: 1px solid #D5DEEA; border-radius: 6px; padding: 7px; selection-background-color: #D9E9FF; }
            QPushButton { background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 7px 12px; min-height: 20px; }
            QPushButton:hover { background: #EFF6FF; border-color: #8CB8F2; }
            QPushButton:disabled { color: #A0AAB8; background: #EEF1F5; }
            QPushButton#primaryButton { background: #1769E0; color: white; border: none; font-weight: 700; padding: 10px 20px; }
            QPushButton#primaryButton:hover { background: #0F57C5; }
            QPushButton#primaryButton:disabled { background: #AABBD3; color: #F7FAFC; }
            QPushButton#navButton { border: none; background: transparent; padding: 7px 14px; color: #52657F; }
            QHeaderView::section { background: #EDF3FA; border: none; border-bottom: 1px solid #DCE4F0; padding: 7px; font-weight: 600; }
            QProgressBar { background: #E8EEF6; border: none; border-radius: 5px; height: 12px; text-align: center; }
            QProgressBar::chunk { background: #2674E8; border-radius: 5px; }
        """)

    def _connect_signals(self):
        self.source_browse_btn.clicked.connect(self._choose_source_dir)
        self.template_browse_btn.clicked.connect(self._choose_template)
        self.source_edit.textChanged.connect(self._on_source_text_changed)
        self.template_edit.textChanged.connect(self._on_any_path_changed)
        self.file_combo.currentTextChanged.connect(self._on_file_selected)
        self.field_edit.textChanged.connect(self._update_start_button)
        self.save_mapping_btn.clicked.connect(self._save_current_mapping)
        self.reset_mapping_btn.clicked.connect(self._reset_current_mapping)
        self.delete_mapping_btn.clicked.connect(self._delete_selected_mapping)
        self.select_all_btn.clicked.connect(self._toggle_select_all)
        self.locate_btn.clicked.connect(self._locate_unmatched)
        self.mapping_table.cellClicked.connect(self._on_mapping_row_clicked)
        self.clear_log_btn.clicked.connect(self.log_view.clear)
        # 【BACKEND BIND】开始处理按钮 -> 后端启动槽函数
        self.start_btn.clicked.connect(self.on_start_process_clicked)
        self.auto_open_check.stateChanged.connect(self._save_auto_open)

    def _load_saved_paths(self):
        self.source_edit.setText(self.user_config.get("last_source_dir", ""))
        self.template_edit.setText(self.user_config.get("last_template", ""))

    def _choose_source_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择源文件目录", self.source_edit.text())
        if path:
            self.source_edit.setText(path)

    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择输出模板", self.template_edit.text(),
            "Excel 工作簿 (*.xlsx)"
        )
        if path:
            self.template_edit.setText(path)

    def _on_source_text_changed(self):
        self._refresh_file_list()
        self._on_any_path_changed()

    def _on_any_path_changed(self):
        self.save_config()
        self._update_start_button()

    def _refresh_file_list(self):
        source = Path(self.source_edit.text().strip()) if self.source_edit.text().strip() else None
        names = []
        if source and source.is_dir():
            names = sorted(
                p.name for p in source.iterdir()
                if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
            )
        current = self.file_combo.currentText()
        self.file_combo.blockSignals(True)
        self.file_combo.clear()
        self.file_combo.addItems(names)
        if current in names:
            self.file_combo.setCurrentText(current)
        self.file_combo.blockSignals(False)

        if names:
            self.file_count_label.setText(f"已发现 {len(names)} 个 .xlsx 文件")
            if self.file_combo.currentText():
                self._load_mapping_for_current_file()
        else:
            self.file_count_label.setText("未发现 .xlsx 文件")
            self.field_edit.clear()
        self._refresh_mapping_table()
        self._update_start_button()

    def _on_file_selected(self, _name):
        self._load_mapping_for_current_file()
        self._update_start_button()

    def _load_mapping_for_current_file(self):
        name = self.file_combo.currentText()
        key = extract_doc_key(name) if name else ""
        self.field_edit.blockSignals(True)
        self.field_edit.setText(self.field_mapping.get(key, ""))
        self.field_edit.blockSignals(False)
        self._update_start_button()

    def _save_current_mapping(self):
        name = self.file_combo.currentText()
        fields = self.field_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "无法保存", "请先选择源文件。")
            return
        if not fields:
            QMessageBox.warning(self, "无法保存", "统计字段不能为空。")
            return
        key = extract_doc_key(name)
        self.field_mapping[key] = fields
        self.save_config()
        self._refresh_mapping_table()
        self.append_log(f"[配置] 已保存单据 [{key}] 的字段配置（{name}）。")
        self._update_start_button()

    def _reset_current_mapping(self):
        name = self.file_combo.currentText()
        if not name:
            return
        key = extract_doc_key(name)
        default_value = self.default_mapping.get(key, "")
        self.field_edit.setText(default_value)
        if default_value:
            self.field_mapping[key] = default_value
        else:
            self.field_mapping.pop(key, None)
        self.save_config()
        self._refresh_mapping_table()
        self.append_log(f"[配置] 已恢复单据 [{key}] 的默认字段配置。")

    def _delete_selected_mapping(self):
        keys_to_delete = []
        for row in range(self.mapping_table.rowCount()):
            check_item = self.mapping_table.item(row, 0)
            key_item = self.mapping_table.item(row, 3)
            if check_item and key_item and check_item.checkState() == Qt.Checked:
                keys_to_delete.append(key_item.text())
        if not keys_to_delete:
            QMessageBox.information(self, "提示", "请先勾选要删除配置的文件。")
            return
        for key in keys_to_delete:
            self.field_mapping.pop(key, None)
        self.save_config()
        self._refresh_mapping_table()
        self.append_log(f"[配置] 已删除 {len(keys_to_delete)} 个单据的字段配置：{', '.join(keys_to_delete)}")
        self._update_start_button()

    def _on_mapping_row_clicked(self, row, _column):
        item = self.mapping_table.item(row, 2)
        if item:
            self.file_combo.setCurrentText(item.text())

    def _locate_unmatched(self):
        total = self.mapping_table.rowCount()
        if total == 0:
            return
        start = getattr(self, "_last_locate_row", -1) + 1
        for i in range(total):
            row = (start + i) % total
            fields_item = self.mapping_table.item(row, 4)
            if not fields_item or not fields_item.text().strip():
                self.mapping_table.selectRow(row)
                self.mapping_table.scrollToItem(fields_item)
                name_item = self.mapping_table.item(row, 2)
                if name_item:
                    self.file_combo.setCurrentText(name_item.text())
                self._last_locate_row = row
                self.append_log(f"[定位] 未匹配文件：{name_item.text() if name_item else ''}")
                return
        QMessageBox.information(self, "定位完成", "所有文件都已配置统计字段。")
        self._last_locate_row = -1

    def _toggle_select_all(self):
        """点一下全选，再点一下全不选，按钮文字跟着变"""
        self._all_checked = not self._all_checked
        state = Qt.Checked if self._all_checked else Qt.Unchecked
        self.select_all_btn.setText("全不选" if self._all_checked else "全选")
        for row in range(self.mapping_table.rowCount()):
            item = self.mapping_table.item(row, 0)
            if item:
                item.setCheckState(state)

    def _get_checked_files(self) -> list:
        """返回表格中勾选的文件名列表"""
        result = []
        for row in range(self.mapping_table.rowCount()):
            check_item = self.mapping_table.item(row, 0)
            name_item = self.mapping_table.item(row, 2)
            if check_item and name_item and check_item.checkState() == Qt.Checked:
                result.append(name_item.text())
        return result

    def _refresh_mapping_table(self):
        names = []
        source_text = self.source_edit.text().strip()
        if source_text:
            source = Path(source_text)
            if source.is_dir():
                names = sorted(
                    p.name for p in source.iterdir()
                    if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
                )
        yellow = QColor("#FFF8E1")
        matched = 0
        self.mapping_table.setRowCount(len(names))
        for row, name in enumerate(names):
            key = extract_doc_key(name)
            fields = self.field_mapping.get(key, "")
            if str(fields).strip():
                matched += 1
            # 第0列：复选框，默认不勾选
            check_item = QTableWidgetItem()
            check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            check_item.setCheckState(Qt.Unchecked)
            self.mapping_table.setItem(row, 0, check_item)
            # 第1列：序号
            self.mapping_table.setItem(row, 1, QTableWidgetItem(str(row + 1)))
            self.mapping_table.setItem(row, 2, QTableWidgetItem(name))
            self.mapping_table.setItem(row, 3, QTableWidgetItem(key))
            fields_item = QTableWidgetItem(str(fields))
            self.mapping_table.setItem(row, 4, fields_item)
            if not str(fields).strip():
                for col in range(5):
                    self.mapping_table.item(row, col).setBackground(yellow)
        total = len(names)
        self.stats_label.setText(
            f"全部文件配置：共 {total} 个文件，已匹配 {matched} 个，未匹配 {total - matched} 个"
        )
        # 刷新后重置全选按钮状态
        self._all_checked = False
        self.select_all_btn.setText("全选")

    def _save_auto_open(self):
        self.save_config()

    def _start_placeholder(self):
        # 已废弃，保留以防外部调用；实际开始按钮已绑定 on_start_process_clicked
        pass

    # ====================== BACKEND SLOT START ======================
    def on_start_process_clicked(self):
        """【后端入口】点击开始处理：启动子线程跑Excel处理，UI不卡死。"""
        if self.process_worker and self.process_worker.isRunning():
            QMessageBox.information(self, "提示", "正在处理中，请等待任务结束")
            return

        source_dir = self.source_edit.text().strip()
        template = self.template_edit.text().strip()
        if not source_dir or not template:
            QMessageBox.warning(self, "提示", "请先选好源目录、输出模板")
            return
        # 输出目录自动创建在源目录下的"已生成"子文件夹
        output_dir = str(Path(source_dir) / "已生成")
        configured = [v for v in self.field_mapping.values() if str(v).strip()]
        if not configured:
            QMessageBox.warning(self, "提示", "至少配置一条单据的统计字段")
            return
        selected_files = self._get_checked_files()
        if not selected_files:
            QMessageBox.warning(self, "提示", "请至少勾选一个要处理的文件")
            return

        # 重置UI状态
        self.log_view.clear()
        self.progress_bar.setValue(0)
        self.success_label.setText("成功：0")
        self.failed_label.setText("失败：0")

        self.process_worker = ExcelProcessWorker(
            source_dir=source_dir,
            output_template=template,
            output_dir=output_dir,
            field_mapping=self.field_mapping,
            selected_files=selected_files,
        )
        self.process_worker.sig_log.connect(self._on_backend_log)
        self.process_worker.sig_progress.connect(self.progress_bar.setValue)
        self.process_worker.sig_count.connect(self._on_backend_count)
        self.process_worker.sig_finish.connect(self._on_backend_finished)
        self.process_worker.start()

    def _on_backend_log(self, text: str):
        """接收后端日志，追加到日志框"""
        self.append_log(text)

    def _on_backend_count(self, success: int, fail: int):
        """接收后端成功/失败计数"""
        self.success_label.setText(f"成功：{success}")
        self.failed_label.setText(f"失败：{fail}")

    def _on_backend_finished(self, ok: bool):
        """后端全部处理完"""
        if ok:
            QMessageBox.information(self, "完成", "全部文件处理结束，请查看日志")
            # 勾选了"处理完成后自动打开输出目录"
            if self.auto_open_check.isChecked():
                source_dir = self.source_edit.text().strip()
                if source_dir:
                    import os
                    out_dir = str(Path(source_dir) / "已生成")
                    os.startfile(out_dir)
        else:
            QMessageBox.warning(self, "结束", "任务异常或被终止，请查看日志")
        self.process_worker = None
    # ====================== BACKEND SLOT END ======================

    def _is_valid_ready(self) -> tuple[bool, str]:
        source_text = self.source_edit.text().strip()
        template_text = self.template_edit.text().strip()

        if not source_text:
            return False, "源文件目录未选择"
        source = Path(source_text)
        if not source.is_dir():
            return False, "源文件目录不存在"
        if not any(p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$") for p in source.iterdir()):
            return False, "源目录中没有 .xlsx 文件"
        if not template_text:
            return False, "输出模板未选择"
        template = Path(template_text)
        if not template.is_file() or template.suffix.lower() != ".xlsx":
            return False, "输出模板不存在或不是 .xlsx 文件"

        # 只要至少有一个单据配置了统计字段即可；未配置的文件处理时自动跳过。
        configured = [k for k, v in self.field_mapping.items() if str(v).strip()]
        if not configured:
            return False, "尚未配置任何单据的统计字段"
        return True, f"配置完整，可开始处理（已配置 {len(configured)} 个单据）"

    def _update_start_button(self):
        ready, reason = self._is_valid_ready()
        self.start_btn.setEnabled(ready)
        self.start_btn.setToolTip(reason)
        if ready:
            self.status_label.setText("配置完成，可以开始处理")
        else:
            self.status_label.setText(f"等待配置：{reason}")

    def append_log(self, message: str, level: str = "info"):
        import re
        timestamp = datetime.now().strftime("%H:%M:%S")
        # 全部用appendHtml，避免appendHtml/appendPlainText混用导致格式继承错乱
        # 先做HTML转义
        safe = message.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        # 换行转<br>，空格用pre样式保留
        safe = safe.replace("\n", "<br>")

        if message.startswith("[成功]"):
            color, weight = "#1E8449", "600"
        elif message.startswith(("[失败]", "[异常]", "[错误]", "[警告]")):
            color, weight = "#C0392B", "600"
        elif message.startswith("[跳过]"):
            color, weight = "#B8860B", "600"
        elif message.startswith("=====") and ("成功" in message or "失败" in message):
            safe = re.sub(r'(成功\s*\d+)', r'<span style="color:#1E8449;font-weight:600">\1</span>', safe)
            safe = re.sub(r'(失败\s*\d+)', r'<span style="color:#C0392B;font-weight:600">\1</span>', safe)
            safe = re.sub(r'(未配置统计字段\s*\d+)', r'<span style="color:#B8860B;font-weight:600">\1</span>', safe)
            color, weight = "#243247", "400"
        else:
            color, weight = "#243247", "400"

        self.log_view.appendHtml(
            f'<div style="white-space:pre-wrap;color:#718096">{timestamp}</div>'
            f'<div style="white-space:pre-wrap;color:{color};font-weight:{weight}">{safe}</div>'
        )

    def closeEvent(self, event):
        self.save_config()
        super().closeEvent(event)
