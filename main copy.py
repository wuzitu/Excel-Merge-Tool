import sys
import os
import time
import shutil
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QProgressBar, QPlainTextEdit, QComboBox, QInputDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QCursor
from config_manager import ConfigManager
from excel_processor import ExcelProcessor


class ExcelMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel合并工具 - 新华三（2025最终修复版）")
        self.resize(1200, 700)

        self.src_dir = ""
        self.out_dir = os.path.join(os.getcwd(), "out_put")
        os.makedirs(self.out_dir, exist_ok=True)

        self.logs_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(self.logs_dir, exist_ok=True)
        self.log_file = os.path.join(self.logs_dir, "debug.log")
        self.log_buffer = []
        self.history_stack = []

        self.config_mgr = ConfigManager()
        self.config = self.config_mgr.config

        layout_main = QHBoxLayout()

        # ===== 左栏 =====
        layout_left = QVBoxLayout()

        # 配置文件管理
        config_layout = QHBoxLayout()
        self.config_selector = QComboBox()
        self.config_selector.addItems(self.config_mgr.list_configs())
        self.config_selector.currentTextChanged.connect(self.change_config)
        config_layout.addWidget(self.config_selector)

        self.btn_save_config = self.create_button("保存配置", "RoyalBlue", self.save_config)
        config_layout.addWidget(self.btn_save_config)

        btn_rename = self.create_button("重命名", "LightSlateGray", self.rename_config_action)
        config_layout.addWidget(btn_rename)

        btn_import = self.create_button("导入", "LightSlateGray", self.import_config_action)
        config_layout.addWidget(btn_import)

        btn_export = self.create_button("导出", "LightSlateGray", self.export_config_action)
        config_layout.addWidget(btn_export)

        layout_left.addLayout(config_layout)

        # 源目录选择
        src_layout = QHBoxLayout()
        self.btn_src = self.create_button("选择源目录", "orange", self.choose_src_dir)
        src_layout.addWidget(self.btn_src)
        self.label_src_path = QLabel("未选择源目录")
        src_layout.addWidget(self.label_src_path)
        layout_left.addLayout(src_layout)

        # 输出目录选择
        out_layout = QHBoxLayout()
        self.btn_out = self.create_button("选择输出目录", "orange", self.choose_out_dir)
        out_layout.addWidget(self.btn_out)
        self.label_out_path = QLabel(f"默认输出: {self.out_dir}")
        out_layout.addWidget(self.label_out_path)
        layout_left.addLayout(out_layout)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["表头名", "单元格位置"])
        layout_left.addWidget(self.table)
        self.load_table_from_config()

        # 表格操作按钮
        btn_table_ops = QHBoxLayout()
        self.btn_add_row = self.create_button("增加行", "DodgerBlue", self.add_row)
        btn_table_ops.addWidget(self.btn_add_row)
        self.btn_delete_row = self.create_button("删除最后一行", "lightcoral", self.delete_last_row)
        btn_table_ops.addWidget(self.btn_delete_row)
        self.btn_undo = self.create_button("撤销", "gray", self.undo_action)
        btn_table_ops.addWidget(self.btn_undo)
        layout_left.addLayout(btn_table_ops)

        # 开始合并按钮
        self.btn_merge = self.create_button("开始合并", "green", self.run_merge)
        layout_left.addWidget(self.btn_merge)

        layout_main.addLayout(layout_left, 3)

        # ===== 右栏 =====
        layout_right = QVBoxLayout()
        progress_layout = QHBoxLayout()
        self.progress = QProgressBar()
        progress_layout.addWidget(self.progress)
        self.label_percent = QLabel("0%")
        progress_layout.addWidget(self.label_percent)
        layout_right.addLayout(progress_layout)

        self.label_status = QLabel("等待操作...")
        self.label_status.setWordWrap(True)
        layout_right.addWidget(self.label_status)

        self.debug_output = QPlainTextEdit()
        self.debug_output.setReadOnly(True)
        layout_right.addWidget(self.debug_output, stretch=1)

        layout_main.addLayout(layout_right, 2)
        self.setLayout(layout_main)

    # 按钮样式（缩小版）
    def btn_style(self, color):
        return f"""
        QPushButton {{
            background-color: {color};
            color: white;
            font-size: 14px;
            font-weight: bold;
            font-family: 'Microsoft YaHei';
            border-radius: 4px;
            height: 34px;
        }}
        QPushButton:hover {{
            background-color: {color};
            opacity: 0.85;
        }}
        QPushButton:pressed {{
            background-color: {color};
            opacity: 0.7;
        }}
        """

    def create_button(self, text, color, func):
        btn = QPushButton(text)
        btn.setCursor(QCursor(Qt.PointingHandCursor))
        btn.setStyleSheet(self.btn_style(color))
        btn.clicked.connect(func)
        return btn

    # 重命名（防空、防已存在）
    def rename_config_action(self):
        old_name = self.config_selector.currentText()
        base_name = os.path.splitext(old_name)[0]
        new_name, ok = QInputDialog.getText(self, "重命名配置文件", "请输入新的配置名称（不含后缀）:", text=base_name)
        if ok and new_name.strip():
            new_name = f"{new_name.strip()}.json"
            old_path = os.path.join(self.config_mgr.configs_dir, old_name)
            new_path = os.path.join(self.config_mgr.configs_dir, new_name)
            if os.path.exists(new_path):
                QMessageBox.warning(self, "提示", f"配置文件 {new_name} 已存在，无法重命名！")
                return
            os.rename(old_path, new_path)
            self.config_selector.blockSignals(True)
            self.config_selector.clear()
            self.config_selector.addItems(self.config_mgr.list_configs())
            self.config_selector.setCurrentText(new_name)
            self.config_selector.blockSignals(False)
            self.change_config(new_name)
            self.log(f"配置文件 {old_name} 重命名为 {new_name}")

    # 导入（防同文件、防已存在）
    def import_config_action(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择配置文件", "", "JSON Files (*.json)")
        if file_path:
            dest_path = os.path.join(self.config_mgr.configs_dir, os.path.basename(file_path))
            if os.path.abspath(file_path) == os.path.abspath(dest_path):
                QMessageBox.information(self, "提示", "该文件已在 configs 目录中，无需导入。")
                return
            if os.path.exists(dest_path):
                QMessageBox.warning(self, "提示", f"配置文件 {os.path.basename(file_path)} 已存在，无法导入！")
                return
            shutil.copy(file_path, dest_path)
            self.config_selector.blockSignals(True)
            self.config_selector.clear()
            self.config_selector.addItems(self.config_mgr.list_configs())
            self.config_selector.setCurrentText(os.path.basename(file_path))
            self.config_selector.blockSignals(False)
            self.change_config(os.path.basename(file_path))
            self.log(f"导入配置文件: {file_path}")

    # 导出
    def export_config_action(self):
        filename = self.config_selector.currentText()
        export_path, _ = QFileDialog.getSaveFileName(self, "导出配置文件", filename, "JSON Files (*.json)")
        if export_path:
            self.config_mgr.export_config(filename, export_path)
            self.log(f"导出配置文件: {filename} 到 {export_path}")

    # 日志
    def log(self, msg):
        ts_msg = f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}"
        self.log_buffer.append(ts_msg)
        if len(self.log_buffer) > 1000:
            self.log_buffer.pop(0)
        self.debug_output.setPlainText("\n".join(self.log_buffer))
        self.debug_output.verticalScrollBar().setValue(self.debug_output.verticalScrollBar().maximum())
        existing_logs = []
        if os.path.exists(self.log_file):
            with open(self.log_file, "r", encoding="utf-8") as f:
                existing_logs = f.readlines()
        existing_logs.append(ts_msg + "\n")
        if len(existing_logs) > 1000:
            existing_logs = existing_logs[-1000:]
        with open(self.log_file, "w", encoding="utf-8") as f:
            f.writelines(existing_logs)

    # 切换配置（加安全判断）
    def change_config(self, filename):
        if not filename:
            return
        path = os.path.join(self.config_mgr.configs_dir, filename)
        if not os.path.exists(path):
            self.log(f"切换配置失败：{path} 不存在")
            return
        self.config = self.config_mgr.load_config(path)
        self.load_table_from_config()
        self.log(f"切换配置文件: {filename}")

    def choose_src_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择源目录")
        if dir_path:
            self.src_dir = dir_path
            self.label_src_path.setText(dir_path)
            self.log(f"选择源目录: {dir_path}")

    def choose_out_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.out_dir = dir_path
            self.label_out_path.setText(dir_path)
            self.log(f"选择输出目录: {dir_path}")
        else:
            self.label_out_path.setText(f"默认输出: {self.out_dir}")
            self.log("输出目录未选择，使用默认 out_put")

    def load_table_from_config(self):
        self.table.setRowCount(0)
        for h in self.config["headers"]:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(h["name"]))
            self.table.setItem(row, 1, QTableWidgetItem(h["cell"]))

    def add_row(self):
        self.table.insertRow(self.table.rowCount())
        self.history_stack.append(("add", None))
        self.log("增加一行空记录")

    def delete_last_row(self):
        last_row = self.table.rowCount() - 1
        if last_row >= 0:
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(last_row, col)
                row_data.append(item.text() if item else "")
            self.history_stack.append(("delete", row_data))
            self.table.removeRow(last_row)
            self.log(f"删除最后一行：第 {last_row+1} 行")
        else:
            QMessageBox.warning(self, "提示", "当前没有行可删除")

    def undo_action(self):
        if not self.history_stack:
            QMessageBox.information(self, "提示", "没有可撤销的操作")
            return
        action, data = self.history_stack.pop()
        if action == "delete":
            row = self.table.rowCount()
            self.table.insertRow(row)
            for col, val in enumerate(data):
                self.table.setItem(row, col, QTableWidgetItem(val))
            self.log("撤销删除操作，恢复一行")
        elif action == "add":
            last_row = self.table.rowCount() - 1
            if last_row >= 0:
                self.table.removeRow(last_row)
            self.log("撤销增加行操作，移除最后一行")

    def save_config(self):
        headers = []
        for row in range(self.table.rowCount()):
            name_item = self.table.item(row, 0)
            cell_item = self.table.item(row, 1)
            if name_item and cell_item:
                headers.append({"name": name_item.text(), "cell": cell_item.text()})
        filename = self.config_selector.currentText()
        path = os.path.join(self.config_mgr.configs_dir, filename)
        self.config_mgr.save_config({"headers": headers}, path)
        self.config = self.config_mgr.config
        self.log(f"保存配置: {filename}")

    def run_merge(self):
        self.save_config()
        if not self.src_dir:
            QMessageBox.warning(self, "提示", "请选择源目录")
            return
        self.progress.setValue(0)
        self.label_percent.setText("0%")
        processor = ExcelProcessor(self.src_dir, self.out_dir, self.config, logger=self.log)
        start_time = time.time()
        try:
            out_path, total_rows = processor.merge_excels(
                progress_callback=lambda v: (self.progress.setValue(v), self.label_percent.setText(f"{v}%")),
                status_callback=lambda s: self.label_status.setText(s)
            )
            elapsed = round(time.time() - start_time, 2)
            QMessageBox.information(self, "合并完成", f"已生成文件：{out_path}\n总记录数：{total_rows}\n耗时：{elapsed}秒")
            os.startfile(self.out_dir)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            self.log(f"合并出错: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    win = ExcelMergerApp()
    win.show()
    app.exec_()