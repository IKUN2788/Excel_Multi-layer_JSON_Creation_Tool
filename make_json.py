import sys
import os
import json
import time
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QFileDialog, QComboBox, QTableWidget, 
                             QTableWidgetItem, QHeaderView, QRadioButton, QButtonGroup, 
                             QMessageBox, QGroupBox, QSplitter, QTextEdit)
from PyQt5.QtCore import Qt
import python_calamine

class JsonConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("通用多层JSON制作工具")
        self.resize(1200, 800)
        self.file_path = ""
        self.sheet_names = []
        self.headers = []
        
        self.init_ui()
        
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 1. 文件选择区域
        file_group = QGroupBox("1. 文件选择")
        file_layout = QHBoxLayout()
        
        self.path_label = QLabel("未选择文件")
        self.path_label.setStyleSheet("border: 1px solid #ccc; padding: 5px; background-color: #f9f9f9;")
        select_btn = QPushButton("选择Excel文件")
        select_btn.clicked.connect(self.load_file)
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.setPlaceholderText("选择工作表")
        self.sheet_combo.currentIndexChanged.connect(self.load_headers)
        
        file_layout.addWidget(self.path_label)
        file_layout.addWidget(select_btn)
        file_layout.addWidget(QLabel("工作表:"))
        file_layout.addWidget(self.sheet_combo)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # 内容区域 (分割线)
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)
        
        # 左侧：第一层键配置 (映射配置)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_group = QGroupBox("2. 第一层键配置 (分类映射)")
        left_inner_layout = QVBoxLayout()
        
        # 选择列
        col_select_layout = QHBoxLayout()
        col_select_layout.addWidget(QLabel("选择列:"))
        self.first_key_combo = QComboBox()
        col_select_layout.addWidget(self.first_key_combo)
        left_inner_layout.addLayout(col_select_layout)
        
        # 映射表
        left_inner_layout.addWidget(QLabel("映射规则 (目标名称 -> 正则表达式):"))
        self.mapping_table = QTableWidget(0, 2)
        self.mapping_table.setHorizontalHeaderLabels(["目标名称 (Key)", "包含的字段 (支持正则)"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        left_inner_layout.addWidget(self.mapping_table)
        
        # 表格操作按钮
        btn_layout = QHBoxLayout()
        add_row_btn = QPushButton("添加规则")
        add_row_btn.clicked.connect(self.add_mapping_row)
        del_row_btn = QPushButton("删除选中规则")
        del_row_btn.clicked.connect(self.delete_mapping_row)
        btn_layout.addWidget(add_row_btn)
        btn_layout.addWidget(del_row_btn)
        left_inner_layout.addLayout(btn_layout)
        
        left_group.setLayout(left_inner_layout)
        left_layout.addWidget(left_group)
        splitter.addWidget(left_widget)
        
        # 右侧：第二层键与值配置
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 第二层键
        second_key_group = QGroupBox("3. 第二层键配置 (唯一标识)")
        second_key_layout = QVBoxLayout()
        h_layout_2 = QHBoxLayout()
        h_layout_2.addWidget(QLabel("选择列:"))
        self.second_key_combo = QComboBox()
        h_layout_2.addWidget(self.second_key_combo)
        second_key_layout.addLayout(h_layout_2)
        second_key_group.setLayout(second_key_layout)
        right_layout.addWidget(second_key_group)
        
        # 值配置
        value_group = QGroupBox("4. 值配置")
        value_layout = QVBoxLayout()
        h_layout_3 = QHBoxLayout()
        h_layout_3.addWidget(QLabel("选择列:"))
        self.value_combo = QComboBox()
        h_layout_3.addWidget(self.value_combo)
        value_layout.addLayout(h_layout_3)
        
        value_layout.addWidget(QLabel("处理模式:"))
        self.mode_group = QButtonGroup(self)
        self.rb_keep = QRadioButton("保留一项 (覆盖)")
        self.rb_accumulate = QRadioButton("累加 (数值求和)")
        self.rb_keep.setChecked(True)
        self.mode_group.addButton(self.rb_keep)
        self.mode_group.addButton(self.rb_accumulate)
        
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(self.rb_keep)
        mode_layout.addWidget(self.rb_accumulate)
        value_layout.addLayout(mode_layout)
        
        value_group.setLayout(value_layout)
        right_layout.addWidget(value_group)
        
        # 预留空间
        right_layout.addStretch()
        splitter.addWidget(right_widget)
        
        # 5. 执行区域
        action_group = QGroupBox("5. 生成")
        action_layout = QVBoxLayout()
        
        run_btn = QPushButton("生成 JSON")
        run_btn.setStyleSheet("font-size: 16px; font-weight: bold; padding: 10px; background-color: #4CAF50; color: white;")
        run_btn.clicked.connect(self.generate_json)
        action_layout.addWidget(run_btn)
        
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(150)
        action_layout.addWidget(self.log_area)
        
        action_group.setLayout(action_layout)
        main_layout.addWidget(action_group)

    def log(self, message):
        self.log_area.append(message)
        QApplication.processEvents()

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.file_path = file_path
            self.path_label.setText(file_path)
            self.log(f"已选择文件: {file_path}")
            
            try:
                with open(self.file_path, 'rb') as f:
                    xls = python_calamine.CalamineWorkbook.from_filelike(f)
                    self.sheet_names = xls.sheet_names
                    self.sheet_combo.clear()
                    self.sheet_combo.addItems(self.sheet_names)
                    
                    # 尝试自动选择 '明细'
                    if '明细' in self.sheet_names:
                        self.sheet_combo.setCurrentText('明细')
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取文件失败: {str(e)}")

    def load_headers(self):
        if not self.file_path or not self.sheet_combo.currentText():
            return
            
        sheet_name = self.sheet_combo.currentText()
        try:
            with open(self.file_path, 'rb') as f:
                xls = python_calamine.CalamineWorkbook.from_filelike(f)
                sheet = xls.get_sheet_by_name(sheet_name)
                rows = iter(sheet.to_python())
                try:
                    self.headers = list(map(str, next(rows)))
                    self.update_combos()
                    self.log(f"已加载工作表 '{sheet_name}' 的表头，共 {len(self.headers)} 列")
                except StopIteration:
                    self.log("工作表为空")
        except Exception as e:
            self.log(f"读取表头失败: {e}")

    def update_combos(self):
        for combo in [self.first_key_combo, self.second_key_combo, self.value_combo]:
            current = combo.currentText()
            combo.clear()
            combo.addItems(self.headers)
            if current in self.headers:
                combo.setCurrentText(current)
            # 尝试智能匹配默认值
            elif combo == self.first_key_combo and '增值费用' in self.headers:
                combo.setCurrentText('增值费用')
            elif combo == self.second_key_combo and '运单号码' in self.headers:
                combo.setCurrentText('运单号码')
            elif combo == self.value_combo and '应付金额' in self.headers:
                combo.setCurrentText('应付金额')

    def add_mapping_row(self):
        row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(row)
        self.mapping_table.setItem(row, 0, QTableWidgetItem("分类名称"))
        self.mapping_table.setItem(row, 1, QTableWidgetItem("匹配正则 (如: 快递|运费)"))

    def delete_mapping_row(self):
        current_row = self.mapping_table.currentRow()
        if current_row >= 0:
            self.mapping_table.removeRow(current_row)

    def generate_json(self):
        if not self.file_path:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return
            
        # 获取配置
        first_key_col = self.first_key_combo.currentText()
        second_key_col = self.second_key_combo.currentText()
        value_col = self.value_combo.currentText()
        
        if not all([first_key_col, second_key_col, value_col]):
            QMessageBox.warning(self, "提示", "请确保所有列选项都已选择")
            return
            
        # 获取映射规则并预编译正则
        mappings = [] # list of (target_name, compiled_regex)
        for row in range(self.mapping_table.rowCount()):
            target_item = self.mapping_table.item(row, 0)
            pattern_item = self.mapping_table.item(row, 1)
            
            if not target_item or not pattern_item:
                continue
                
            target = target_item.text().strip()
            pattern_str = pattern_item.text().strip()
            
            if target and pattern_str:
                try:
                    # 预编译正则，忽略大小写
                    regex = re.compile(pattern_str, re.IGNORECASE)
                    mappings.append((target, regex))
                except re.error as e:
                    self.log(f"错误: 正则表达式 '{pattern_str}' 无效: {e}")
                    QMessageBox.warning(self, "错误", f"规则 '{target}' 的正则表达式无效:\n{e}")
                    return
        
        if not mappings:
            QMessageBox.warning(self, "提示", "请至少添加一个有效的映射规则")
            return
            
        is_accumulate = self.rb_accumulate.isChecked()
        
        # 初始化结果字典
        result_data = {m[0]: {} for m in mappings}
        
        try:
            t1 = time.time()
            with open(self.file_path, 'rb') as f:
                xls = python_calamine.CalamineWorkbook.from_filelike(f)
                sheet = xls.get_sheet_by_name(self.sheet_combo.currentText())
                rows = iter(sheet.to_python())
                
                # 获取索引
                head = list(map(str, next(rows)))
                try:
                    idx_first = head.index(first_key_col)
                    idx_second = head.index(second_key_col)
                    idx_value = head.index(value_col)
                except ValueError as e:
                    self.log(f"列名匹配失败: {e}")
                    return

                processed_count = 0
                match_count = 0
                
                for row in rows:
                    if len(row) <= max(idx_first, idx_second, idx_value):
                        continue
                        
                    raw_first = row[idx_first]
                    raw_second = row[idx_second]
                    raw_value = row[idx_value]
                    
                    # 匹配第一层键
                    target_key = None
                    # 转换为字符串进行匹配
                    check_val = str(raw_first) if raw_first is not None else ""
                    
                    for target, regex in mappings:
                        if regex.search(check_val):
                            target_key = target
                            break
                    
                    if not target_key:
                        continue
                        
                    # 处理值
                    val = 0
                    if is_accumulate:
                        try:
                            val = float(raw_value) if raw_value is not None else 0.0
                        except:
                            val = 0.0
                    else:
                        val = str(raw_value) if raw_value is not None else ""
                    
                    second_key = str(raw_second) if raw_second is not None else ""
                    
                    if is_accumulate:
                        current = result_data[target_key].get(second_key, 0.0)
                        result_data[target_key][second_key] = current + val
                    else:
                        result_data[target_key][second_key] = val
                        
                    match_count += 1
                    processed_count += 1

            # 保存文件
            # output_dir = os.path.join(os.path.dirname(self.file_path), "json数据")
            # 修改为保存到根目录 (当前工作目录)
            output_dir = os.path.join(os.getcwd(), "json数据")
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, "generated_data.json")
            
            # 如果是累加模式，可能需要格式化一下小数位
            if is_accumulate:
                 for k1 in result_data:
                     for k2 in result_data[k1]:
                         result_data[k1][k2] = round(result_data[k1][k2], 2)

            with open(output_path, "w", encoding="utf-8") as f_w:
                json.dump(result_data, f_w, ensure_ascii=False, indent=2)
                
            t2 = time.time()
            self.log(f"处理完成! 耗时: {t2-t1:.2f}秒")
            self.log(f"共匹配数据: {match_count} 条")
            self.log(f"文件已保存至: {output_path}")
            
            QMessageBox.information(self, "成功", f"JSON生成成功！\n保存路径: {output_path}")
            
        except Exception as e:
            self.log(f"处理出错: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"处理过程出错: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = JsonConverterApp()
    window.show()
    sys.exit(app.exec_())
