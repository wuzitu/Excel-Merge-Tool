import os
import time
import pandas as pd
import openpyxl
from datetime import datetime

class ExcelProcessor:
    def __init__(self, src_dir, out_dir, config, logger=None):
        self.src_dir = src_dir
        self.out_dir = out_dir
        self.config = config
        self.logger = logger if logger else print

    def merge_excels(self, progress_callback=None, status_callback=None):
        start_time = time.time()

        file_list = [f for f in os.listdir(self.src_dir) if f.lower().endswith((".xlsx", ".xls"))]
        total_files = len(file_list)
        if total_files == 0:
            raise ValueError("源目录没有可处理的Excel文件")

        merged_rows = []
        processed_files = 0
        header_names = [h["name"] for h in self.config["headers"]]

        for file_name in file_list:
            file_path = os.path.join(self.src_dir, file_name)
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                row_data = []
                for h in self.config["headers"]:
                    cell_value = ws[h["cell"]].value
                    row_data.append(cell_value)
                merged_rows.append(row_data)
            except Exception as e:
                self.logger(f"文件 {file_name} 处理失败: {str(e)}")
                continue

            processed_files += 1
            percent = int(processed_files / total_files * 100)
            if progress_callback:
                progress_callback(percent)

            elapsed = time.time() - start_time
            avg_time = elapsed / processed_files
            remain_secs = (total_files - processed_files) * avg_time

            if status_callback:
                status_callback(file_name, processed_files, total_files, remain_secs)

        # 生成 DataFrame
        result_df = pd.DataFrame(merged_rows, columns=header_names)

        # ===================== 文件命名规则 =====================
        # 1. 获取源目录最后一级文件夹名
        folder_name = os.path.basename(os.path.normpath(self.src_dir))
        # 2. 数据行数
        row_count = len(result_df)
        # 3. 组合基础文件名
        base_filename = f"{folder_name}_集合_{row_count}条.xlsx"
        out_path = os.path.join(self.out_dir, base_filename)
        # 4. 检查是否冲突（已有同名文件），冲突则加时间后缀
        if os.path.exists(out_path):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_filename = f"{folder_name}_集合_{row_count}条_{timestamp}.xlsx"
            out_path = os.path.join(self.out_dir, base_filename)
        # =======================================================

        result_df.to_excel(out_path, index=False)
        return out_path, len(result_df)