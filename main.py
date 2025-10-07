import sys
import os
import time
import shutil
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QProgressBar, QPlainTextEdit, QComboBox, QInputDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QCursor, QColor, QIcon
from config_manager import ConfigManager
from excel_processor import ExcelProcessor


class ExcelMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel合并工具 - 兔子版v1.0")
        self.resize(1400, 700)
        
        # 设置窗口图标
        icon_path = os.path.join(os.getcwd(), 'my_icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

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

        # 左栏布局
        layout_left = QVBoxLayout()

        # 第一行：仅下拉框
        config_selector_layout = QHBoxLayout()
        self.config_selector = QComboBox()
        self.config_selector.setMinimumWidth(400)
        self.config_selector.setFont(QFont("Microsoft YaHei", 10))
        self.config_selector.addItems(self.config_mgr.list_configs())
        self.config_selector.currentTextChanged.connect(self.change_config)
        config_selector_layout.addWidget(QLabel("选择配置文件："))
        config_selector_layout.addWidget(self.config_selector)
        layout_left.addLayout(config_selector_layout)

        # 第二行：配置按钮
        config_buttons_layout = QHBoxLayout()
        config_buttons_layout.addWidget(self.create_button("保存配置", "RoyalBlue", self.save_config))
        config_buttons_layout.addWidget(self.create_button("新增配置", "Teal", self.add_new_config))
        config_buttons_layout.addWidget(self.create_button("刷新列表", "DarkGreen", self.refresh_config_selector))
        config_buttons_layout.addWidget(self.create_button("重命名", "LightSlateGray", self.rename_config_action))
        config_buttons_layout.addWidget(self.create_button("导入", "LightSlateGray", self.import_config_action))
        config_buttons_layout.addWidget(self.create_button("导出", "LightSlateGray", self.export_config_action))
        config_buttons_layout.addWidget(self.create_button("配置目录", "DarkSlateBlue", self.open_config_dir))
        layout_left.addLayout(config_buttons_layout)

        # 源目录
        src_layout = QHBoxLayout()
        src_layout.addWidget(self.create_button("选择源目录", "orange", self.choose_src_dir))
        self.label_src_path = QLabel("未选择源目录")
        src_layout.addWidget(self.label_src_path)
        layout_left.addLayout(src_layout)

        # 输出目录
        out_layout = QHBoxLayout()
        out_layout.addWidget(self.create_button("选择输出目录", "orange", self.choose_out_dir))
        self.label_out_path = QLabel(f"默认输出: {self.out_dir}")
        out_layout.addWidget(self.label_out_path)
        layout_left.addLayout(out_layout)

        # 表格 headers
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["表头名", "单元格位置"])
        layout_left.addWidget(self.table)
        self.load_table_from_config()

        # 表格操作按钮
        table_ops = QHBoxLayout()
        table_ops.addWidget(self.create_button("增加行", "DodgerBlue", self.add_row))
        table_ops.addWidget(self.create_button("删除最后一行", "lightcoral", self.delete_last_row))
        table_ops.addWidget(self.create_button("撤销", "gray", self.undo_action))
        layout_left.addLayout(table_ops)

        # 合并按钮
        layout_left.addWidget(self.create_button("开始合并", "green", self.run_merge))
        layout_main.addLayout(layout_left, 2)

        # 右栏布局
        layout_right = QVBoxLayout()
        prog_layout = QHBoxLayout()
        self.progress = QProgressBar()
        prog_layout.addWidget(self.progress)
        self.label_percent = QLabel("0%")
        prog_layout.addWidget(self.label_percent)
        layout_right.addLayout(prog_layout)

        self.label_status = QLabel("等待操作...")
        self.label_status.setWordWrap(True)
        layout_right.addWidget(self.label_status)

        self.debug_output = QPlainTextEdit()
        self.debug_output.setReadOnly(True)
        layout_right.addWidget(self.debug_output, stretch=1)

        layout_main.addLayout(layout_right, 4)
        self.setLayout(layout_main)

        self.load_last_selected_config()

    def btn_style(self, color):
        base = QColor(color)
        hover = QColor(min(base.red() + 20, 255),
                       min(base.green() + 20, 255),
                       min(base.blue() + 20, 255))
        pressed = QColor(max(base.red() - 20, 0),
                         max(base.green() - 20, 0),
                         max(base.blue() - 20, 0))
        return f"""
        QPushButton {{
            background-color: rgb({base.red()}, {base.green()}, {base.blue()});
            color: white;
            font-size: 14px;
            font-weight: bold;
            font-family: 'Microsoft YaHei';
            border-radius: 4px;
            height: 34px;
        }}
        QPushButton:hover {{
            background-color: rgb({hover.red()}, {hover.green()}, {hover.blue()});
        }}
        QPushButton:pressed {{
            background-color: rgb({pressed.red()}, {pressed.green()}, {pressed.blue()});
            padding-left: 2px;
            padding-top: 2px;
        }}
        """

    def create_button(self, text, color, func):
        btn = QPushButton(text)
        btn.setCursor(QCursor(Qt.PointingHandCursor))
        btn.setStyleSheet(self.btn_style(color))
        btn.clicked.connect(func)
        return btn

    def rename_config_action(self):
        old_name = self.config_selector.currentText()
        base_name = os.path.splitext(old_name)[0]
        new_name, ok = QInputDialog.getText(self, "重命名配置文件", "请输入新的配置名称（不含后缀）:", text=base_name)
        if ok and new_name.strip():
            new_name = f"{new_name.strip()}.json"
            old_path = os.path.join(self.config_mgr.configs_dir, old_name)
            new_path = os.path.join(self.config_mgr.configs_dir, new_name)
            if os.path.exists(new_path):
                QMessageBox.warning(self, "提示", f"配置文件 {new_name} 已存在")
                return
            os.rename(old_path, new_path)
            self.refresh_config_list(new_name)
            self.log(f"配置文件 {old_name} 重命名为 {new_name}")

    def import_config_action(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择配置文件", "", "JSON Files (*.json)")
        if file_path:
            dest_path = os.path.join(self.config_mgr.configs_dir, os.path.basename(file_path))
            if os.path.abspath(file_path) == os.path.abspath(dest_path):
                QMessageBox.information(self, "提示", "该文件已在目录中")
                return
            if os.path.exists(dest_path):
                QMessageBox.warning(self, "提示", "同名文件已存在")
                return
            shutil.copy(file_path, dest_path)
            self.refresh_config_list(os.path.basename(file_path))
            self.log(f"导入配置文件: {file_path}")

    def export_config_action(self):
        filename = self.config_selector.currentText()
        export_path, _ = QFileDialog.getSaveFileName(self, "导出配置文件", filename, "JSON Files (*.json)")
        if export_path:
            self.config_mgr.export_config(filename, export_path)
            self.log(f"导出配置文件 {filename} 到 {export_path}")

    def open_config_dir(self):
        os.startfile(self.config_mgr.configs_dir)
        self.log(f"打开配置目录: {self.config_mgr.configs_dir}")

    def refresh_config_list(self, current_name):
        self.config_selector.blockSignals(True)
        self.config_selector.clear()
        self.config_selector.addItems(self.config_mgr.list_configs())
        self.config_selector.setCurrentText(current_name)
        self.config_selector.blockSignals(False)
        self.change_config(current_name)

    def load_last_selected_config(self):
        last_file = os.path.join(self.logs_dir, "last_config.txt")
        if os.path.exists(last_file):
            with open(last_file, "r", encoding="utf-8") as f:
                last_name = f.read().strip()
            if last_name in self.config_mgr.list_configs():
                self.config_selector.setCurrentText(last_name)
                self.change_config(last_name)

    def save_last_selected_config(self, filename):
        with open(os.path.join(self.logs_dir, "last_config.txt"), "w", encoding="utf-8") as f:
            f.write(filename)

    def change_config(self, filename):
        if not filename:
            return
        path = os.path.join(self.config_mgr.configs_dir, filename)
        if not os.path.exists(path):
            return
        self.config = self.config_mgr.load_config(path)
        self.load_table_from_config()
        self.save_last_selected_config(filename)

    def log(self, msg):
        ts_msg = f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}"
        self.log_buffer.append(ts_msg)
        if len(self.log_buffer) > 1000:
            self.log_buffer.pop(0)
        self.debug_output.setPlainText("\n".join(self.log_buffer))
        self.debug_output.verticalScrollBar().setValue(self.debug_output.verticalScrollBar().maximum())
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(ts_msg + "\n")

    def update_status_log(self, file_name, processed_files, total_files, remaining_seconds):
        remaining_time = time.strftime('%H:%M:%S', time.gmtime(remaining_seconds))
        info = (f"文件: {file_name}\n"
                f"已统计文件数量: {processed_files}/{total_files}\n"
                f"预计剩余时间: {remaining_time}")
        self.label_status.setText(info)
        self.log(info)

    def load_table_from_config(self):
        self.table.setRowCount(0)
        for h in self.config["headers"]:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(h["name"]))
            self.table.setItem(row, 1, QTableWidgetItem(h["cell"]))

    def add_row(self):
        # 在添加行之前保存当前状态到历史堆栈
        self.save_current_state()
        self.table.insertRow(self.table.rowCount())

    def delete_last_row(self):
        last_row = self.table.rowCount() - 1
        if last_row >= 0:
            # 在删除行之前保存当前状态到历史堆栈
            self.save_current_state()
            self.table.removeRow(last_row)

    def save_current_state(self):
        # 保存当前表格状态到历史堆栈
        headers = []
        for row in range(self.table.rowCount()):
            name_item = self.table.item(row, 0)
            cell_item = self.table.item(row, 1)
            if name_item and cell_item:
                headers.append({"name": name_item.text(), "cell": cell_item.text()})
        self.history_stack.append(headers)
        # 限制历史堆栈大小，避免内存占用过大
        if len(self.history_stack) > 50:
            self.history_stack.pop(0)

    def undo_action(self):
        # 实现撤销功能 - 撤销上一次操作（新增或删除行）
        if self.history_stack:
            # 移除当前状态的保存，直接恢复到上一个状态
            headers = self.history_stack.pop()
            # 清空表格并恢复历史状态
            self.table.setRowCount(0)
            for h in headers:
                row = self.table.rowCount()
                self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(h["name"]))
                self.table.setItem(row, 1, QTableWidgetItem(h["cell"]))
            self.log("已执行撤销操作")
        else:
            QMessageBox.information(self, "提示", "没有可撤销的操作")

    def refresh_config_selector(self):
        # 刷新配置文件列表，用于感知手动添加的配置文件
        current_config = self.config_selector.currentText()
        self.refresh_config_list(current_config)
        self.log("配置文件列表已刷新")

    def add_new_config(self):
        # 新增配置文件功能
        new_name, ok = QInputDialog.getText(self, "新增配置文件", "请输入新的配置名称（不含后缀）:")
        if ok and new_name.strip():
            new_name = f"{new_name.strip()}.json"
            new_path = os.path.join(self.config_mgr.configs_dir, new_name)
            if os.path.exists(new_path):
                QMessageBox.warning(self, "提示", f"配置文件 {new_name} 已存在")
                return
            # 创建新的配置文件，使用空的headers列表
            self.config_mgr.save_config({"headers": []}, new_path)
            # 刷新配置列表并选中新创建的配置
            self.refresh_config_list(new_name)
            self.log(f"新增配置文件: {new_name}")

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
        # 保存后刷新配置列表
        self.refresh_config_list(filename)
        self.log(f"配置已保存到: {filename}")

    def choose_src_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择源目录")
        if d:
            self.src_dir = d
            self.label_src_path.setText(d)

    def choose_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.out_dir = d
            self.label_out_path.setText(d)
        else:
            self.label_out_path.setText(f"默认输出: {self.out_dir}")

    def run_merge(self):
        self.save_config()
        if not self.src_dir:
            QMessageBox.warning(self, "提示", "请选择源目录")
            return
        self.progress.setValue(0)
        self.label_percent.setText("0%")
        processor = ExcelProcessor(self.src_dir, self.out_dir, self.config, logger=self.log)
        try:
            out_path, total_rows = processor.merge_excels(
                progress_callback=lambda v: (self.progress.setValue(v),
                                             self.label_percent.setText(f"{v}%")),
                status_callback=lambda f, pf, tf, r: self.update_status_log(f, pf, tf, r)
            )
            QMessageBox.information(self, "合并完成", f"文件：{out_path}\n记录数：{total_rows}")
            os.startfile(self.out_dir)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            self.log(f"合并出错: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    
    # 设置应用程序图标，确保在Windows任务栏上显示
    icon_path = os.path.join(os.getcwd(), 'my_icon.ico')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    win = ExcelMergerApp()
    win.show()
    app.exec_()