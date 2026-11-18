# -*- coding: utf-8 -*-
# ============================================================
# 📌【BACKEND IMPORTS START】
# 后端业务线程依赖，UI层不要在这里import任何Qt控件
# ============================================================
import os
import re
from pathlib import Path

from PySide6.QtCore import QThread, Signal
# ============================================================
# 📌【BACKEND IMPORTS END】
# ============================================================


# ============================================================
# 📌【BACKEND START】
# 后端工作线程：遍历文件、调度单文件处理、回传状态给UI
# 业务逻辑统一写在 handle_single_file() 里，不要操作Qt控件
# ============================================================
class ExcelProcessWorker(QThread):
    # ---- 回传给UI的信号 ----
    sig_log = Signal(str)            # 一行日志
    sig_progress = Signal(int)        # 进度 0~100
    sig_count = Signal(int, int)      # (成功数, 失败数)
    sig_finish = Signal(bool)         # 全部结束（是否正常）

    def __init__(self, source_dir: str, output_template: str,
                 output_dir: str, field_mapping: dict,
                 selected_files: list = None, parent=None):
        super().__init__(parent)
        self.source_dir = source_dir
        self.output_template = output_template
        self.output_dir = output_dir
        self.field_mapping = field_mapping or {}
        self.selected_files = set(selected_files or [])
        self._is_running = True

    def stop_task(self):
        """UI层调用，请求终止任务"""
        self._is_running = False

    # 单据名提取规则必须和 ui/main_window.py 里的 extract_doc_key 保持一致
    @staticmethod
    def extract_doc_key(filename: str) -> str:
        stem = Path(filename).stem
        chinese = "".join(re.findall(r"[\u4e00-\u9fff]", stem))
        chinese = chinese.replace("月", "")
        return chinese if chinese else stem

    def run(self):
        self._is_running = True
        success_cnt = 0
        fail_cnt = 0

        # 确保输出目录存在（自动创建在源目录下的"已生成"子文件夹）
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)

        src = Path(self.source_dir)
        all_files = [
            p.name for p in src.iterdir()
            if p.is_file() and p.suffix.lower() in (".xlsx", ".xls") and not p.name.startswith("~$")
        ]
        # 只处理勾选的文件
        if self.selected_files:
            file_list = sorted(f for f in all_files if f in self.selected_files)
        else:
            file_list = sorted(all_files)
        skipped_unchecked = len(all_files) - len(file_list)

        total = len(file_list)
        if total == 0:
            self.sig_log.emit("[警告] 没有勾选要处理的文件")
            self.sig_finish.emit(False)
            return

        self.sig_log.emit(f"===== 开始处理，共 {total} 个文件（未勾选排除 {skipped_unchecked} 个）=====")
        skip_cnt = 0

        # 预加载模板，拿到所有sheet名，用于提前检查
        try:
            from openpyxl import load_workbook
            wb_tmp = load_workbook(self.output_template, read_only=True)
            template_sheets = set(wb_tmp.sheetnames)
            wb_tmp.close()
            self.sig_log.emit(f"[信息] 模板已加载，共 {len(template_sheets)} 个sheet")
        except Exception as e:
            self.sig_log.emit(f"[错误] 模板加载失败：{e}")
            self.sig_finish.emit(False)
            return

        for idx, name in enumerate(file_list):
            if not self._is_running:
                self.sig_log.emit("[信息] 任务被手动终止")
                break

            self.sig_progress.emit(int((idx / total) * 100))
            full_path = str(src / name)
            doc_key = self.extract_doc_key(name)
            field_str = self.field_mapping.get(doc_key, "")
            fields = [f.strip() for f in field_str.split("|") if f.strip()] if field_str else []

            # 没匹配到字段的文件：不进业务逻辑，不计成功/失败，只数个数
            if not fields:
                skip_cnt += 1
                continue

            # 模板里没有对应sheet：算失败，不读Excel，省时间
            if doc_key not in template_sheets:
                self.sig_log.emit(f"[失败] {name}：模板中找不到sheet「{doc_key}」")
                fail_cnt += 1
                self.sig_count.emit(success_cnt, fail_cnt)
                continue

            self.sig_log.emit(f"--- [{idx + 1}/{total}] {name}（单据：{doc_key}）")
            try:
                # ================================================
                # ✅ 业务入口：在这里写单文件读取、预处理、提取、输出
                # ================================================
                ok = self.handle_single_file(
                    file_path=full_path,
                    doc_key=doc_key,
                    target_fields=fields,
                )
                # ================================================
                if ok:
                    success_cnt += 1
                    self.sig_log.emit(f"[成功] {name}")
                else:
                    fail_cnt += 1
                    self.sig_log.emit(f"[失败] {name} 业务返回失败")
            except Exception as e:
                fail_cnt += 1
                self.sig_log.emit(f"[异常] {name}：{e}")

            self.sig_count.emit(success_cnt, fail_cnt)

        self.sig_progress.emit(100)
        self.sig_log.emit(
            f"===== 处理结束：成功 {success_cnt}，失败 {fail_cnt}，未配置统计字段 {skip_cnt} ====="
        )
        self.sig_finish.emit(True)

    # ============================================================
    # ✅ 业务开发核心位置（Excel逻辑）
    # 根据 doc_key 区分不同单据的起始行/起始单元格；
    # 读取、预处理、按 target_fields 提取、套用模板输出到 output_dir。
    # 需要打日志就 self.sig_log.emit("...")，不要直接碰Qt控件。
    # 返回 True 表示成功，False 表示失败。
    # ============================================================
    def handle_single_file(self, file_path: str, doc_key: str, target_fields: list) -> bool:
        import pandas as pd

        # 起始行
        if doc_key == '完工入库成本查询':
            header_row = 3
        elif doc_key == '折旧明细表':
            header_row = 4
        else:
            header_row = 1

        # 读取excel
        if '物料编码' in target_fields:
            df = pd.read_excel(file_path, skiprows=header_row - 1, dtype={'物料编码': str})
        else:
            df = pd.read_excel(file_path, skiprows=header_row - 1)

        # 清除所有小计和总计
        df = df[~df.iloc[:, 0].str.contains('小计|合计', na=False)]

        # 只保留目标字段列
        df = df[target_fields]

        # 空行处理
        df = df.dropna(how="all")

        if doc_key == '完工入库成本查询':
            # 起始为第3行 共4列(产品名称|基本单位|数量|金额)，前2列空值填补，最后1列需检查总计，按最后2列统计

            # 空值填充
            cols = target_fields[:2]
            df.loc[:, cols] = df.loc[:, cols].ffill()
    
            # 统计
            df = df.groupby(target_fields[:2]).sum().reset_index()
    
            # 添加合计行
            sums = df.iloc[:, 2:].sum()
            total_row = pd.DataFrame([['合计', '', *sums]], columns=target_fields)
            df = pd.concat([df, total_row], ignore_index=True)
        elif doc_key == '折旧明细表':
            # 起始为第4行
            # 资产名称|卡片编码|资产数量（卡片）|期初原值|原值增加|原值减少|本期折旧额|本年累计折旧额|期末原值|期末累计折旧|期末净值

            # 前2列空值填补
            cols_1 = target_fields[:2]
            df.loc[:, cols_1] = df.loc[:, cols_1].ffill()
            # 其余所有数值列作空值补0
            df.iloc[:, 2:] = df.iloc[:, 2:].fillna(0)

            # 统计
            df = df.groupby(target_fields[:2]).sum().reset_index()

            # 添加合计行
            sums = df.iloc[:, 2:].sum()
            total_row = pd.DataFrame([['合计', '', *sums]], columns=target_fields)
            df = pd.concat([df, total_row], ignore_index=True)
        elif doc_key in ('其他出库单', '其他入库单', '销售出库单', '销售退货单', '成本调整单'):
            # 起始正常 共5列(申请类型|物料编码|物料名称|单位|实发数量)，前4列空值填补，最后1列需检查总计，按最后1列统计

            # 空值填充
            cols = target_fields[:4]
            df.loc[:, cols] = df.loc[:, cols].ffill()
            # 统计
            df = df.groupby(target_fields[:4]).sum().reset_index()

            # 添加合计行
            sums = df.iloc[:, 4:].sum()
            total_row = pd.DataFrame([['合计', '', '', '', *sums]], columns=target_fields)
            df = pd.concat([df, total_row], ignore_index=True)
        else:
            # 起始正常 共4列(供应商|物料名称|库存单位|实退数量)，前3列空值填补，最后1列需检查总计，按最后1列统计

            # 空值填充
            cols = target_fields[:3]
            df.loc[:, cols] = df.loc[:, cols].ffill()

            # 统计
            df = df.groupby(target_fields[:3]).sum().reset_index()

            # 添加合计行
            sums = df.iloc[:, 3:].sum()
            total_row = pd.DataFrame([['合计', '', '', *sums]], columns=target_fields)
            df = pd.concat([df, total_row], ignore_index=True)

        if doc_key == '成本调整单':
            # 统计完需重新调整列名顺序

            df = df.reindex(columns=['物料编码', '物料名称', '调整金额', '调整原因', '仓库']) # 重新排列

        self.sig_log.emit(f"  单据={doc_key}，表头在第{header_row}行，共{len(df)}行数据")
        # self.sig_log.emit(f"  单据={doc_key}，表头在第{header_row}行，共{len(df)}行数据, 数据:\n{df}\n")

        # ================================================
        # ✅ 写入模板并另存为新文件
        # ================================================
        from core.excel_writer import write_df_to_template
        out_path = write_df_to_template(
            template_path=self.output_template,
            output_dir=self.output_dir,
            source_filename=os.path.basename(file_path),
            doc_key=doc_key,
            df=df,
        )
        self.sig_log.emit(f"  已输出到：{out_path}")
        return True
# ============================================================
# 📌【BACKEND END】
# ============================================================
