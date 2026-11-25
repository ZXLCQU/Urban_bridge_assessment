import sys
import os
import json
import time
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QLabel, QTableWidget, QTableWidgetItem, QHeaderView,
                               QComboBox, QCheckBox, QHBoxLayout, QTabWidget, QPushButton, QSizePolicy, QRadioButton,
                               QMessageBox, QLineEdit, QDialogButtonBox, QDialog, QFileDialog)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QFont
import copy
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

class BridgeComponentBase(QMainWindow):
    """桥梁构件评估基类"""
    def __init__(self):
        super().__init__()
        self.component_type = ""  # 构件类型(桥面系/上部结构等)
        self.table = None
        self.deduct_points = {}
        self.load_config()
    def load_config(self):
        """加载配置文件"""
        # 创建basic_data目录（如果不存在）
        basic_data_dir = os.path.join(os.path.dirname(__file__), "basic_data")
        os.makedirs(basic_data_dir, exist_ok=True)
        config_path = os.path.join(basic_data_dir, "Solid_belly_arch_bridge.json")
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
    def init_base_ui(self):
        """初始化基础UI"""
        widget = QWidget()
        self.setCentralWidget(widget)
        layout = QVBoxLayout(widget)
        # 创建标题
        title = QLabel(f"{self.component_type}各构件评分等级、扣分表")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                padding: 4px 8px;
                margin: 1px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                          stop:0 #f6f7fa, stop:1 #e6e6e6);
                border-radius: 2px;
                color: #333333;
            }
        """)
        title.setFixedHeight(30)
        layout.addWidget(title)
        # 创建表格
        self.init_table()
        layout.addWidget(self.table)
    def init_table(self):
        """初始化表格"""
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(['部位', '损坏类型', '损坏程度', '扣分值DP(hi)',
                                              '权重ω(hi)', 'DP(hi)*ω(hi)', 'MDP(h)',
                                              '权重ω(h)', 'BSI(m)', 'BCI(m)'])
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.table.horizontalHeader().setFixedHeight(25)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        # 设置表格样式
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 4px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 2px 4px;
            }
        """)
        # 设置列宽和调整模式
        self.table.setColumnWidth(0, 100)
        self.table.setColumnWidth(1, 120)
        self.table.setColumnWidth(2, 120)
        self.table.setColumnWidth(3, 100)
        self.table.setColumnWidth(4, 100)
        self.table.setColumnWidth(5, 100)
        self.table.setColumnWidth(6, 65)
        self.table.setColumnWidth(7, 65)
        self.table.setColumnWidth(8, 65)
        self.table.setColumnWidth(9, 65)
        # 设置每列的调整模式
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(8, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(9, QHeaderView.Fixed)
        # 隐藏垂直表头
        self.table.verticalHeader().setVisible(False)
    def calculate_weight(self, row):
        """计算权重"""
        # 1. 获取当前行的损坏类型
        damage_type = self.table.item(row, 1).text()
        # 2. 获取该损坏类型所属部位的行范围
        start_row, end_row = self.get_part_range(damage_type)
        # 3. 计算该部位的总扣分值
        total_deduct = sum(float(self.table.item(r, 3).text()) for r in range(start_row, end_row))
        # 4. 获取当前行的扣分值
        deduct_points = float(self.table.item(row, 3).text())
        # 5. 计算uhi
        uhi = deduct_points / total_deduct if total_deduct > 0 else 0
        # 6. 计算权重
        weight = round((3.0 * (uhi ** 3) - 5.5 * (uhi ** 2) + 3.5 * uhi), 2)
        # 7. 计算扣分值和权重的乘积
        product = round(deduct_points * weight, 2)
        return weight, product
    def calculate_mdp(self, start_row, end_row):
        """计算构件综合扣分值MDP(h)"""
        # 1. 计算该构件所有行第6列的乘积之和（综合扣分值）
        total_product = sum(float(self.table.item(r, 5).text()) for r in range(start_row, end_row))
        # 2. 获取该构件第4列中的最大扣分值
        max_deduct = max(float(self.table.item(r, 3).text()) for r in range(start_row, end_row))
        # 3. 根据条件判断最终的MDP(h)值
        if total_product <= max_deduct:
            mdp = max_deduct
        else:
            mdp = min(total_product, 100)
        return round(mdp, 2)
    def calculate_bsi_bci(self):
        """计算桥梁结构指数BSI(m)和桥梁状况指数BCI(m)"""
        # 获取所有MDP(h)值
        mdp_values = []
        weighted_sum = 0
        row = 0
        while row < self.table.rowCount():
            if self.table.item(row, 6):  # 检查MDP(h)列
                mdp = float(self.table.item(row, 6).text())
                weight = float(self.table.item(row, 7).text())
                mdp_values.append(mdp)
                weighted_sum += mdp * weight
            row += self.table.rowSpan(row, 0)
        # 计算BSI(m)：100 - max(MDP(h))
        bsi = 100 - max(mdp_values) if mdp_values else 100
        # 计算BCI(m)：100 - Σ(MDP(h) * ω(h))
        bci = 100 - weighted_sum
        # 更新表格
        self.table.item(0, 8).setText(f"{bsi:.2f}")
        self.table.item(0, 9).setText(f"{bci:.2f}")
    def update_deduct_points(self, damage_type, row):
        """更新扣分值、权重、乘积和MDP(h)"""
        combo = self.table.cellWidget(row, 2)  # 获取下拉框
        current_degree = combo.currentText()  # 获取下拉框选择的损坏程度
        # 如果损坏类型和损坏程度都存在于配置文件中，更新扣分值和权重
        if damage_type in self.deduct_points and current_degree in self.deduct_points[damage_type]:
            # 1. 自动匹配扣分值
            points = self.deduct_points[damage_type][current_degree]
            self.table.item(row, 3).setText(f"{float(points):.0f}")
            # 2. 更新当前行的权重和乘积
            weight, product = self.calculate_weight(row)
            self.table.item(row, 4).setText(f"{weight:.2f}")
            self.table.item(row, 5).setText(f"{product:.2f}")
            # 3. 更新同一部位其他行的权重和乘积
            start_row, end_row = self.get_part_range(damage_type)
            for r in range(start_row, end_row):
                if r != row:
                    weight, product = self.calculate_weight(r)
                    self.table.item(r, 4).setText(f"{weight:.2f}")
                    self.table.item(r, 5).setText(f"{product:.2f}")
            # 4. 更新MDP(h)值
            mdp = self.calculate_mdp(start_row, end_row)
            self.table.item(start_row, 6).setText(f"{mdp:.2f}")
            # 5. 更新BSI和BCI值
            self.calculate_bsi_bci()
class DeckSystem(BridgeComponentBase):
    """桥面系评估"""
    def __init__(self):
        super().__init__()
        self.component_type = "桥面系"
        self.init_ui()
    def init_ui(self):
        """初始化桥面系UI"""
        self.init_base_ui()
        self.load_deck_config()
        self.init_deck_data()
    def load_deck_config(self):
        """加载桥面系配置"""
        bridge_deck = self.config["桥面系"]
        for part, damages in bridge_deck.items():
            for damage_type, info in damages.items():
                self.deduct_points[damage_type] = info["scores"]
    def create_component_cell(self, name):
        """创建带复选框的部位单元格"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(5, 0, 5, 0)
        checkbox = QCheckBox(name)
        checkbox.setChecked(True)  # 默认选中
        checkbox.stateChanged.connect(self.calculate_component_weights)
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        return widget
    def calculate_component_weights(self):
        """计算各构件权重"""
        standard_weights = {
            "桥面铺装": 0.30,
            "桥头平顺": 0.15,
            "伸缩缝装置": 0.25,
            "排水系统": 0.10,
            "栏杆或护栏": 0.10,
            "人行道块件": 0.10
        }
        # 检查每个构件是否存在
        existing_components = {}
        total_standard_weight = 0
        # 遍历所有构件，检查复选框状态并设置权重
        row = 0
        while row < self.table.rowCount():
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    component_name = checkbox.text()
                    if not checkbox.isChecked():
                        # 未选中的构件权重设为0
                        self.table.item(row, 7).setText("0.00")
                    elif component_name in standard_weights:
                        # 选中的构件使用标准权重
                        existing_components[component_name] = standard_weights[component_name]
                        total_standard_weight += standard_weights[component_name]
            row += self.table.rowSpan(row, 0)
        # 如果没有选中任何构件，直接返回
        if not existing_components:
            return
        # 重新分配权重
        weights = []
        for component, weight in existing_components.items():
            adjusted_weight = weight / total_standard_weight
            weights.append((component, adjusted_weight))
        # 处理舍入误差
        total_rounded = sum(round(w[1], 2) for w in weights)
        if total_rounded != 1:
            # 找到最大权重的构件
            max_component = max(weights, key=lambda x: x[1])
            # 调整最大权重构件的值
            weights = [(c, round(w, 2) if c != max_component[0] else
            round(w + (1 - total_rounded), 2)) for c, w in weights]
            # 更新表格中的权重值
        row = 0
        while row < self.table.rowCount():
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    component_name = checkbox.text()
                    for name, weight in weights:
                        if name == component_name:
                            self.table.item(row, 7).setText(f"{weight:.2f}")
                            break
            row += self.table.rowSpan(row, 0)
        # 更新BSI和BCI值
        self.calculate_bsi_bci()
    def init_deck_data(self):
        """初始化桥面系数据"""
        # 从配置文件构建数据结构
        data = []
        for part_name, part_data in self.config["桥面系"].items():
            data.append({
                'name': part_name,
                'items': list(part_data.keys())
            })
        # 计算总行数
        total_rows = sum(len(group['items']) for group in data)
        self.table.setRowCount(total_rows)
        # 填充数据并合并单元格
        current_row = 0
        for group in data:
            start_row = current_row
            # 填充损坏类型
            for item in group['items']:
                # 设置损坏类型
                damage_item = QTableWidgetItem(item)
                damage_item.setTextAlignment(Qt.AlignCenter)
                damage_item.setFlags(damage_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(current_row, 1, damage_item)
                # 创建损坏程度下拉框
                combo = QComboBox()
                options = self.config["桥面系"][group['name']][item]["options"]
                combo.addItems(options)
                combo.setEditable(True)
                combo.lineEdit().setAlignment(Qt.AlignCenter)
                combo.lineEdit().setReadOnly(True)
                combo.currentTextChanged.connect(
                    lambda text, damage_type=item, row=current_row:
                    self.update_deduct_points(damage_type, row)
                )
                # 设置下拉框样式
                combo.setStyleSheet("""
                    QComboBox {
                        border: 1px solid #ddd;
                        border-radius: 3px;
                        padding: 1px 4px 1px 3px;
                        background: white;
                        min-width: 100px;
                    }
                    QComboBox::drop-down {
                        border: none;
                        border-left: 1px solid #ddd;
                        width: 20px;
                    }
                    QComboBox::down-arrow {
                        width: 8px;
                        height: 8px;
                        margin-top: 1px;
                        margin-right: 4px;
                        border-top: 2px solid #666;
                        border-right: 2px solid #666;
                        transform: rotate(135deg);
                    }
                    QComboBox QAbstractItemView {
                        border: 1px solid #ddd;
                        selection-background-color: #f0f0f0;
                        min-width: 70px;
                    }
                    QComboBox QAbstractItemView::item {
                        height: 24px;
                        padding: 4px;
                        text-align: center;
                    }
                """)
                self.table.setCellWidget(current_row, 2, combo)
                # 初始化其他单元格
                for col in [3, 4, 5]:
                    item = QTableWidgetItem('0' if col == 3 else '0.00')
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(current_row, col, item)
                current_row += 1
            # 创建部位复选框
            self.table.setCellWidget(start_row, 0, self.create_component_cell(group['name']))
            # 合并单元格
            if len(group['items']) > 1:
                self.table.setSpan(start_row, 0, len(group['items']), 1)
            # 初始化MDP(h)和权重ω(h)单元格
            for col in [6, 7]:
                item = QTableWidgetItem('0.00')
                item.setTextAlignment(Qt.AlignCenter)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(start_row, col, item)
                if len(group['items']) > 1:
                    self.table.setSpan(start_row, col, len(group['items']), 1)
        # 初始化BSI和BCI单元格
        for col in [8, 9]:
            item = QTableWidgetItem('100.00')
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(0, col, item)
            self.table.setSpan(0, col, total_rows, 1)
        # 数据初始化完成后，调用权重计算函数
        self.calculate_component_weights()
    def get_part_range(self, damage_type):
        """获取桥面系部位范围"""
        if damage_type in ["网裂或龟裂", "波浪及车辙", "坑槽", "碎裂或破碎", "坑洞", "桥面贯通横缝", "桥面贯通纵缝"]:
            return 0, 7  # 桥面铺装部位（第1-7行，索引0-6）
        elif damage_type in ["桥头沉降", "台背下沉"]:
            return 7, 9  # 桥头平顺部位（第8-9行，索引7-8）
        elif damage_type in ['螺帽松动', '缝内沉积物阻塞', '止水带破损、老化', '钢材料破损', '接缝处铺装碎边',
                             '接缝处高差', '钢材料翘曲变形', '结构缝宽异常', '伸缩缝处异常声响']:
            return 9, 18  # 伸缩缝装置部位（第10-18行，索引9-17）
        elif damage_type in ["泄水管阻塞", "残缺脱落", "桥面积水", "防水层"]:
            return 18, 22  # 排水系统部位（第19-22行，索引18-21）
        elif damage_type in ["露筋锈蚀", "松动错位", "丢失残缺"]:
            return 22, 25  # 栏杆或护栏部位（第23-25行，索引22-24）
        elif damage_type in ["网裂", "松动或变形", "残缺"]:
            return 25, 28  # 人行道块件部位（第26-28行，索引25-27）
class SuperStructure(BridgeComponentBase):
    """上部结构评估"""
    def __init__(self):
        super().__init__()
        self.component_type = "上部结构"
        self.current_span = 1  # 当前跨号
        self.span_data = {1: {}}  # 存储各跨数据
        self.init_ui()
    def init_ui(self):
        """初始化上部结构UI"""
        # 创建主布局
        widget = QWidget()
        self.setCentralWidget(widget)
        layout = QVBoxLayout(widget)
        layout.setSpacing(0)  # 减小整体间距
        layout.setContentsMargins(5, 5, 5, 5)  # 减小边距
        # 添加跨号控制面板
        span_control = QWidget()
        span_layout = QHBoxLayout(span_control)
        span_layout.setContentsMargins(0, 0, 0, 5)
        add_span_btn = QPushButton("添加新跨")
        add_span_btn.clicked.connect(self.add_span)
        add_span_btn.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
        """)
        self.span_combo = QComboBox()
        self.span_combo.addItem("第1跨")
        self.span_combo.currentIndexChanged.connect(self.switch_span)
        self.span_combo.setStyleSheet("""
            QComboBox {
                border: 1px solid #ddd;
                border-radius: 3px;
                padding: 4px;
                min-width: 100px;
            }
        """)
        del_span_btn = QPushButton("删除当前跨")
        del_span_btn.clicked.connect(self.delete_span)
        del_span_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        span_layout.addWidget(add_span_btn)
        span_layout.addWidget(self.span_combo)
        span_layout.addWidget(del_span_btn)
        span_layout.addStretch()
        layout.addWidget(span_control)
        # 创建标题
        title = QLabel(f"{self.component_type}各构件评分等级、扣分表")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                padding: 4px 8px;
                margin: 1px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                        stop:0 #f6f7fa, stop:1 #e6e6e6);
                border-radius: 2px;
                color: #333333;
            }
        """)
        layout.addWidget(title)
        # 初始化表格
        self.init_table()
        self.table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        layout.addWidget(self.table)
        # 创建汇总表格
        self.summary_table = QTableWidget()
        self.summary_table.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.summary_table.setFixedHeight(250)
        self.summary_table.setColumnCount(4)
        self.summary_table.setHorizontalHeaderLabels(['桥跨序号', 'BCI(si)', 'BCI(s)', 'BSI(s)'])
        # 设置汇总表格样式
        self.summary_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 2px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 2px;
            }
        """)
        # 设置汇总表格属性
        self.summary_table.setColumnWidth(0, 100)
        self.summary_table.setColumnWidth(1, 100)
        self.summary_table.setColumnWidth(2, 100)
        self.summary_table.setColumnWidth(3, 100)
        self.summary_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.summary_table.verticalHeader().setVisible(False)
        self.summary_table.verticalHeader().setDefaultSectionSize(22)  # 设置行高
        # 将表格添加到布局中
        layout.addWidget(self.summary_table)
        # 加载配置和初始化数据
        self.load_super_config()
        self.init_super_data()
    def update_summary(self):
        """更新汇总信息"""
        try:
            # 1. 先移除所有已存在的合并单元格
            self.summary_table.clearSpans()
            # 2. 清空汇总表格
            self.summary_table.setRowCount(0)
            # 3. 获取所有有效的跨号（按顺序）
            valid_spans = sorted(self.span_data.keys())
            if not valid_spans:
                return
            # 4. 计算BCI(si)的平均值和最小值
            bci_values = []
            for span_num in valid_spans:
                bci = self.span_data[span_num].get('bci', 100.00)
                bci_values.append(bci)
            bci_average = sum(bci_values) / len(bci_values) if bci_values else 100.00
            bci_min = min(bci_values) if bci_values else 100.00
            # 5. 添加各跨数据
            for span_num in valid_spans:
                row = self.summary_table.rowCount()
                self.summary_table.insertRow(row)
                # 添加跨号
                span_item = QTableWidgetItem(f"第{span_num}跨")
                span_item.setTextAlignment(Qt.AlignCenter)
                self.summary_table.setItem(row, 0, span_item)
                # 添加BCI(si)值
                bci = self.span_data[span_num].get('bci', 100.00)
                bci_item = QTableWidgetItem(f"{bci:.2f}")
                bci_item.setTextAlignment(Qt.AlignCenter)
                self.summary_table.setItem(row, 1, bci_item)
                # 添加BCI(s)值和BSI(s)值
                if row == 0:
                    # 先创建单元格
                    bci_s_item = QTableWidgetItem(f"{bci_average:.2f}")
                    bci_s_item.setTextAlignment(Qt.AlignCenter)
                    self.summary_table.setItem(row, 2, bci_s_item)
                    bsi_item = QTableWidgetItem(f"{bci_min:.2f}")
                    bsi_item.setTextAlignment(Qt.AlignCenter)
                    self.summary_table.setItem(row, 3, bsi_item)
            # 6. 最后再进行合并单元格操作
            if len(valid_spans) > 1:
                try:
                    self.summary_table.setSpan(0, 2, len(valid_spans), 1)
                    self.summary_table.setSpan(0, 3, len(valid_spans), 1)
                except Exception as e:

                    QMessageBox.warning(self, "合并单元格错误", f"合并单元格时出错: {str(e)}")
        except Exception as e:

            QMessageBox.warning(self, "更新汇总错误", f"更新汇总信息时出错: {str(e)}")
    def init_table(self):
        """初始化表格"""
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        headers = ['部位', '损坏类型', '损坏程度', '扣分值DP(ijk)',
                   '权重ω(ijk)', 'DP(ijk)*ω(ijk)', 'SDP(ij)',
                   '权重ω(ij)', 'BCI(si)']
        self.table.setHorizontalHeaderLabels(headers)
        # 设置表格基本属性
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.table.horizontalHeader().setFixedHeight(25)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.verticalHeader().setVisible(False)
        # 设置表格样式
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 4px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 2px 4px;
            }
            /* 添加以下样式以支持文字自动换行 */
            QTableWidget QTableWidgetItem {
                padding: 2px;
            }
        """)
        # 设置列宽
        column_widths = [107, 127, 127, 107, 107, 107, 75, 75, 76]
        for col, width in enumerate(column_widths):
            self.table.setColumnWidth(col, width)
            # 修改第一列的调整模式为自动换行
            if col == 0:
                self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Fixed)
                # 启用自动换行
                self.table.setWordWrap(True)
            else:
                self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Fixed)
    def init_super_data(self):
        """初始化上部结构数据"""
        # 从配置文件构建数据结构
        data = []
        for part_name, part_data in self.config["上部结构"].items():
            data.append({
                'name': part_name,
                'items': list(part_data.keys())
            })
        # 计算总行数
        total_rows = sum(len(group['items']) for group in data)
        self.table.setRowCount(total_rows)
        # 填充数据并合并单元格
        current_row = 0
        for group in data:
            start_row = current_row
            # 填充损坏类型
            for item in group['items']:
                # 设置损坏类型
                damage_item = QTableWidgetItem(item)
                damage_item.setTextAlignment(Qt.AlignCenter)
                damage_item.setFlags(damage_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(current_row, 1, damage_item)
                # 创建损坏程度下拉框
                combo = QComboBox()
                options = self.config["上部结构"][group['name']][item]["options"]
                combo.addItems(options)
                combo.setEditable(True)
                combo.lineEdit().setAlignment(Qt.AlignCenter)
                combo.lineEdit().setReadOnly(True)
                combo.currentTextChanged.connect(
                    lambda text, damage_type=item, row=current_row:
                    self.update_deduct_points(damage_type, row)
                )
                # 设置下拉框样式
                combo.setStyleSheet("""
                    QComboBox {
                        border: 1px solid #ddd;
                        border-radius: 3px;
                        padding: 1px 4px 1px 3px;
                        background: white;
                        min-width: 100px;
                    }
                    QComboBox::drop-down {
                        border: none;
                        border-left: 1px solid #ddd;
                        width: 20px;
                    }
                    QComboBox::down-arrow {
                        width: 8px;
                        height: 8px;
                        margin-top: 1px;
                        margin-right: 4px;
                        border-top: 2px solid #666;
                        border-right: 2px solid #666;
                        transform: rotate(135deg);
                    }
                    QComboBox QAbstractItemView {
                        border: 1px solid #ddd;
                        selection-background-color: #f0f0f0;
                        min-width: 70px;
                    }
                    QComboBox QAbstractItemView::item {
                        height: 24px;
                        padding: 4px;
                        text-align: center;
                    }
                """)
                self.table.setCellWidget(current_row, 2, combo)
                # 初始化其他单元格
                for col in [3, 4, 5]:
                    item = QTableWidgetItem('0' if col == 3 else '0.00')
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(current_row, col, item)
                current_row += 1
            # 创建部位复选框
            self.table.setCellWidget(start_row, 0, self.create_component_cell(group['name']))
            # 合并单元格
            if len(group['items']) > 1:
                self.table.setSpan(start_row, 0, len(group['items']), 1)
            # 初始化SDP(ij)和权重ω(ij)单元格
            for col in [6, 7]:
                item = QTableWidgetItem('0.00')
                item.setTextAlignment(Qt.AlignCenter)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(start_row, col, item)
                if len(group['items']) > 1:
                    self.table.setSpan(start_row, col, len(group['items']), 1)
        # 初始化BCI单元格（合并整个表格的第9列）
        item = QTableWidgetItem('100.00')
        item.setTextAlignment(Qt.AlignCenter)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(0, 8, item)
        self.table.setSpan(0, 8, total_rows, 1)  # 合并整个表格的BCI列
        # 在初始化完表格数据后，立即计算权重
        self.calculate_component_weights()
        # 保存当前跨数据并更新汇总
        self.save_current_span_data()
        self.update_summary()
    def save_current_span_data(self):
        """保存当前跨的数据"""
        span_data = {
            'table_data': [],
            'bci': float(self.table.item(0, 8).text()),
            'checkbox_states': {}  # 新增：保存复选框状态
        }
        row = 0
        while row < self.table.rowCount():
            # 保存复选框状态
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    span_data['checkbox_states'][row] = checkbox.isChecked()
            # 保存行数据
            row_data = {
                'damage_type': self.table.item(row, 1).text() if self.table.item(row, 1) else "",
                'damage_degree': self.table.cellWidget(row, 2).currentText() if self.table.cellWidget(row, 2) else "",
                'deduct_points': self.table.item(row, 3).text() if self.table.item(row, 3) else "0",
                'weight': self.table.item(row, 4).text() if self.table.item(row, 4) else "0.00",
                'product': self.table.item(row, 5).text() if self.table.item(row, 5) else "0.00",
                'sdp': self.table.item(row, 6).text() if self.table.item(row, 6) else "0.00",
                'component_weight': self.table.item(row, 7).text() if self.table.item(row, 7) else "0.00"
            }
            span_data['table_data'].append(row_data)
            row += 1
        self.span_data[self.current_span] = span_data
    def load_span_data(self, span_number, skip_clear=False):
        """加载指定跨号的数据
        Args:
            span_number: 要加载的跨号
            skip_clear: 是否跳过清空表格数据
        """
        if span_number in self.span_data:
            # 只有在非跳过模式下才清空表格数据
            if not skip_clear:
                self.clear_table_data()
            else:
                # 断开所有下拉框的信号连接
                for row in range(self.table.rowCount()):
                    combo = self.table.cellWidget(row, 2)
                    if combo and isinstance(combo, QComboBox):
                        try:
                            # 保存原始信号连接
                            combo.blockSignals(True)
                            # 重置为默认状态
                            combo.setCurrentIndex(0)
                        except Exception as e:
                            QMessageBox.warning(self, "信号连接错误", f"断开信号连接时出错: {str(e)}")

            span_data = self.span_data[span_number]
            # 恢复复选框状态
            for row, is_checked in span_data.get('checkbox_states', {}).items():
                cell_widget = self.table.cellWidget(row, 0)
                if cell_widget:
                    checkbox = cell_widget.findChild(QCheckBox)
                    if checkbox:
                        checkbox.setChecked(is_checked)
            # 恢复表格数据
            for row, row_data in enumerate(span_data.get('table_data', [])):
                if row < self.table.rowCount():
                    # 设置损坏程度
                    if self.table.cellWidget(row, 2) and 'damage_degree' in row_data:
                        self.table.cellWidget(row, 2).setCurrentText(row_data['damage_degree'])
                    # 恢复其他数值
                    columns = {
                        3: 'deduct_points',
                        4: 'weight',
                        5: 'product',
                        6: 'sdp',
                        7: 'component_weight'
                    }
                    for col, key in columns.items():
                        if row_data.get(key) and self.table.item(row, col):
                            self.table.item(row, col).setText(row_data[key])
            # 恢复BCI值
            if 'bci' in span_data and self.table.item(0, 8):
                self.table.item(0, 8).setText(f"{span_data['bci']:.2f}")
            # 重新连接所有下拉框的信号
            for row in range(self.table.rowCount()):
                combo = self.table.cellWidget(row, 2)
                if combo and isinstance(combo, QComboBox):
                    try:
                        combo.blockSignals(False)
                    except Exception as e:
                        QMessageBox.warning(self, "信号连接错误", f"重新连接信号时出错: {str(e)}")
    def add_span(self):
        """添加新跨"""
        # 保存当前跨的数据
        self.save_current_span_data()
        # 创建新跨
        new_span = len(self.span_data) + 1
        # 初始化新跨的数据结构（不复制之前跨的数据）
        self.span_data[new_span] = {
            'table_data': [],
            'bci': 100.00,
            'checkbox_states': {}
        }
        # 初始化新跨的数据（保持复选框状态，但所有数据从零开始）
        for row in range(self.table.rowCount()):
            # 保存复选框状态
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    self.span_data[new_span]['checkbox_states'][row] = checkbox.isChecked()
            # 初始化行数据（全部置零）
            row_data = {
                'damage_type': self.table.item(row, 1).text() if self.table.item(row, 1) else "",
                'damage_degree': "无",  # 默认损坏程度
                'deduct_points': "0",
                'weight': "0.00",
                'product': "0.00",
                'sdp': "0.00",
                'component_weight': self.table.item(row, 7).text() if self.table.item(row, 7) else "0.00"
            }
            self.span_data[new_span]['table_data'].append(row_data)
        # 更新UI
        self.span_combo.addItem(f"第{new_span}跨")
        self.span_combo.setCurrentIndex(new_span - 1)
        self.clear_table_data()
        self.calculate_component_weights()
        self.update_summary()
    def delete_span(self):
        """删除当前跨"""
        if len(self.span_data) > 1:
            # 先保存当前跨的数据
            self.save_current_span_data()
            # 获取当前跨号和目标跨号
            current_index = self.span_combo.currentIndex()
            target_index = current_index - 1 if current_index > 0 else 0
            target_span = target_index + 1
            # 断开信号连接
            self.span_combo.currentIndexChanged.disconnect(self.switch_span)
            # 删除当前跨的数据并重新编号
            new_span_data = {}
            current_new_span = 1
            for old_span in sorted(self.span_data.keys()):
                if old_span != self.current_span:
                    new_span_data[current_new_span] = copy.deepcopy(self.span_data[old_span])
                    current_new_span += 1
            # 更新数据存储
            self.span_data = new_span_data
            # 重建下拉框选项
            self.span_combo.clear()
            for i in range(1, len(self.span_data) + 1):
                self.span_combo.addItem(f"第{i}跨")
            # 设置新的当前跨号
            self.current_span = target_span
            # 先加载数据
            self.load_span_data(self.current_span)
            # 重新连接信号并设置索引
            self.span_combo.currentIndexChanged.connect(self.switch_span)
            self.span_combo.setCurrentIndex(target_index)
            # 更新汇总
            self.update_summary()
    def switch_span(self, index):
        # 保存当前跨的数据
        self.save_current_span_data()
        # 更新当前跨号
        self.current_span = index + 1
        # 加载目标跨的数据（使用skip_clear=True参数避免清空表格）
        self.load_span_data(self.current_span, skip_clear=True)
        # 重新计算权重和更新汇总
        self.calculate_component_weights()
        self.update_summary()
    def clear_table_data(self):
        """清空表格数据但保持复选框状态"""
        for row in range(self.table.rowCount()):
            # 保持复选框状态不变
            if self.table.cellWidget(row, 2):
                self.table.cellWidget(row, 2).setCurrentIndex(0)
            # 清空数值单元格
            for col in [3, 4, 5, 6]:  # 不清空权重(7)和BCI(8)列
                if self.table.item(row, col):
                    # 扣分值列(3)使用整数0，其他列使用0.00
                    value = "0" if col == 3 else "0.00"
                    self.table.item(row, col).setText(value)
    def calculate_bsi_bci(self):
        """重写计算方法（只计算BCI）"""
        mdp_values = []
        weighted_sum = 0
        row = 0
        while row < self.table.rowCount():
            if self.table.item(row, 6):  # 检查SDP(ij)列
                mdp = float(self.table.item(row, 6).text())
                weight = float(self.table.item(row, 7).text())
                mdp_values.append(mdp)
                weighted_sum += mdp * weight
            row += self.table.rowSpan(row, 0)
        # 只计算BCI值
        bci = 100 - weighted_sum
        self.table.item(0, 8).setText(f"{bci:.2f}")
        # 保存并更新汇总信息
        self.save_current_span_data()
        self.update_summary()
    def load_super_config(self):
        """加载上部结构配置"""
        super_structure = self.config["上部结构"]
        for part, damages in super_structure.items():
            for damage_type, info in damages.items():
                self.deduct_points[damage_type] = info["scores"]
    def get_part_range(self, damage_type):
        """获取上部结构部位范围"""
        if damage_type in ["桥面贯通横缝", "主拱圈变形", "砌体缺损", "砌体脱落", "砂浆松动缺失","表面风化损坏", "表面网状裂缝", "构件断裂", "结构裂缝", "混凝土剥离", "露筋锈蚀", "节点开裂", "裂缝处渗水"]:
            return 0, 13  # 主拱圈构件部位
        elif damage_type in ["表面风化剥落", "侧墙变形", "砌体缺损","砌体脱落", "砂浆松动缺失"]:
            return 13, 18  # 拱上构造
        elif damage_type in ["微弯板网裂", "微弯板剥落露筋", "连接件(杆)开裂","连接件(杆)剥落露筋", "桥面贯通纵缝"]:
            return 18, 23  # 横向联系部位
    def calculate_component_weights(self):
        """计算各构件权重"""
        standard_weights = {
            "主拱圈": 0.50,
            "拱上构造": 0.20,
            "横向联系": 0.30
        }
        # 检查每个构件是否存在
        existing_components = {}
        total_standard_weight = 0
        row = 0
        while row < self.table.rowCount():
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    component_name = checkbox.text()
                    if not checkbox.isChecked():
                        # 未选中的构件权重设为0
                        self.table.item(row, 7).setText("0.00")
                    elif component_name in standard_weights:
                        existing_components[component_name] = standard_weights[component_name]
                        total_standard_weight += standard_weights[component_name]
            row += self.table.rowSpan(row, 0)
        if not existing_components:
            return
        # 重新计算权重
        for row in range(self.table.rowCount()):
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox and checkbox.isChecked():
                    component_name = checkbox.text()
                    if component_name in standard_weights:
                        weight = standard_weights[component_name] / total_standard_weight
                        self.table.item(row, 7).setText(f"{weight:.2f}")
        # 更新计算结果
        self.calculate_bsi_bci()
    def create_component_cell(self, name):
        """创建带复选框的部位单元格"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(5, 2, 5, 2)
        checkbox = QCheckBox(name)
        checkbox.setChecked(True)  # 默认选中
        # 直接连接信号到权重计算函数
        checkbox.stateChanged.connect(lambda: self.calculate_component_weights())
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        return widget
class SubStructure(BridgeComponentBase):
    """下部结构评估"""
    def __init__(self):
        super().__init__()
        self.component_type = "下部结构"
        self.structure_type = "桥台"  # 默认为桥台
        self.current_pier = None  # 初始化时不设置编号
        self.abutment_data = {}  # 桥台数据
        self.pier_data = {}  # 桥墩数据
        self.all_components = []  # 用于记录所有构件的添加顺序和类型
        self.init_ui()

    def init_ui(self):
        """初始化下部结构UI"""
        # 创建主布局
        widget = QWidget()
        self.setCentralWidget(widget)
        main_layout = QVBoxLayout(widget)
        main_layout.setSpacing(5)
        main_layout.setContentsMargins(5, 5, 5, 5)
        # 添加控制面板
        control_panel = QWidget()
        control_layout = QHBoxLayout(control_panel)
        control_layout.setContentsMargins(0, 0, 0, 5)
        # 左侧：结构类型切换
        type_group = QWidget()
        type_layout = QHBoxLayout(type_group)
        type_layout.setContentsMargins(0, 0, 0, 0)
        self.abutment_radio = QRadioButton("桥台")
        self.pier_radio = QRadioButton("桥墩")
        self.abutment_radio.setChecked(True)
        self.abutment_radio.toggled.connect(self.switch_structure_type)
        self.pier_radio.toggled.connect(self.switch_structure_type)
        type_layout.addWidget(self.abutment_radio)
        type_layout.addWidget(self.pier_radio)
        control_layout.addWidget(type_group)
        # 中间：编号选择下拉框
        self.pier_combo = QComboBox()
        self.pier_combo.addItem("请添加")  # 修改：初始显示"请添加"
        self.pier_combo.setMinimumWidth(100)
        self.pier_combo.currentIndexChanged.connect(self.switch_pier)
        control_layout.addWidget(self.pier_combo)
        # 右侧：添加/删除按钮
        btn_group = QWidget()
        btn_layout = QHBoxLayout(btn_group)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        add_btn = QPushButton("添加")
        self.del_btn = QPushButton("删除")  # 修改：保存为实例变量
        add_btn.clicked.connect(self.add_pier)
        self.del_btn.clicked.connect(self.delete_pier)  # 修改：使用实例变量
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(self.del_btn)
        control_layout.addWidget(btn_group)
        control_layout.addStretch()
        main_layout.addWidget(control_panel)
        # 添加标题
        title = QLabel(f"{self.component_type}各构件评分等级、扣分表")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                padding: 4px 8px;
                margin: 1px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                        stop:0 #f6f7fa, stop:1 #e6e6e6);
                border-radius: 2px;
                color: #333333;
            }
        """)
        main_layout.addWidget(title)
        # 初始化主表格
        self.init_table()
        main_layout.addWidget(self.table)
        # 创建汇总表格
        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(4)
        self.summary_table.setHorizontalHeaderLabels(['桥台/墩号', 'BCI(xj)', 'BCI(x)', 'BSI(x)'])
        self.summary_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 2px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 2px;
            }
        """)
        # 设置汇总表格属性
        self.summary_table.setColumnWidth(0, 100)
        self.summary_table.setColumnWidth(1, 100)
        self.summary_table.setColumnWidth(2, 100)
        self.summary_table.setColumnWidth(3, 100)
        self.summary_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.summary_table.verticalHeader().setVisible(False)
        self.summary_table.verticalHeader().setDefaultSectionSize(22)
        self.summary_table.setFixedHeight(220)  # 设置固定高度
        # 直接将汇总表格添加到主布局
        main_layout.addWidget(self.summary_table)
        # 加载配置和初始化数据
        self.load_sub_config()
        self.init_sub_data()
        self.load_structure_data()
    def add_pier(self):
        """添加新的桥台/墩"""
        # 获取所有已存在的构件编号（包括桥台和桥墩）
        all_numbers = []
        for comp in self.all_components:
            if comp['number'] is not None:
                all_numbers.append(comp['number'])
        # 确定新编号
        if not all_numbers:
            new_pier = 0
        else:
            new_pier = max(all_numbers) + 1
        # 保存当前数据（如果有）- 移到这里，确保在创建新构件前保存
        if self.current_pier is not None and self.current_pier >= 0:
            self.save_current_data()
        # 初始化新数据
        new_data = {
            'type': self.structure_type,
            'table_data': [],
            'bci': 100.00,
            'checkbox_states': {}
        }
        # 记录新构件
        self.all_components.append({
            'number': new_pier,
            'type': self.structure_type
        })
        # 保存到对应的数据结构
        if self.structure_type == "桥台":
            self.abutment_data[new_pier] = new_data
        else:
            self.pier_data[new_pier] = new_data
        # 更新下拉框
        self.update_pier_combo()
        # 设置新添加的构件为当前选中项
        index = self.pier_combo.findText(f"第{new_pier}#{'台' if self.structure_type == '桥台' else '墩'}")
        if index >= 0:
            self.pier_combo.setCurrentIndex(index)
        # 清空表格数据
        self.clear_table_data()
        # 更新汇总
        self.update_summary()
    def delete_pier(self):
        """删除当前构件"""
        # 检查是否有选中的构件
        if self.pier_combo.currentText() == "请添加":
            return
        # 保存当前编号以供后续使用
        pier_to_delete = self.current_pier
        # 从数据结构中删除
        if self.structure_type == "桥台":
            if pier_to_delete in self.abutment_data:
                del self.abutment_data[pier_to_delete]
        else:
            if pier_to_delete in self.pier_data:
                del self.pier_data[pier_to_delete]
        # 从构件列表中删除
        self.all_components = [comp for comp in self.all_components
                            if not (comp['type'] == self.structure_type and
                                    comp['number'] == pier_to_delete)]
        # 重新排序所有构件编号
        self.reorder_components(pier_to_delete)
        # 清空当前构件编号
        self.current_pier = None
        # 清空汇总表格
        self.summary_table.setRowCount(0)
        # 断开下拉框信号，避免触发不必要的数据保存
        self.pier_combo.currentIndexChanged.disconnect(self.switch_pier)
        # 获取当前类型的所有构件编号
        current_data = self.abutment_data if self.structure_type == "桥台" else self.pier_data
        valid_keys = [k for k in current_data.keys() if k is not None and isinstance(k, (int, float))]
        if valid_keys:
            # 更新下拉框
            self.update_pier_combo()
            # 修改：优先选择比被删除编号小1的构件
            target_pier = pier_to_delete - 1
            # 如果目标编号不存在，则找到小于被删除编号的最大编号
            if target_pier not in valid_keys:
                prev_keys = [k for k in valid_keys if k < pier_to_delete]
                if prev_keys:
                    # 如果有更小的编号，选择最大的那个
                    target_pier = max(prev_keys)
                else:
                    # 如果没有更小的编号，选择最小的编号
                    target_pier = min(valid_keys)
            self.current_pier = target_pier
            # 在下拉框中选择对应的项
            index = self.pier_combo.findText(f"第{target_pier}#{'台' if self.structure_type == '桥台' else '墩'}")
            if index >= 0:
                self.pier_combo.setCurrentIndex(index)
            # 加载目标构件的数据
            self.load_structure_data()
        else:
            # 如果没有其他构件，重置界面
            self.pier_combo.clear()
            self.pier_combo.addItem("请添加")
            self.clear_table_data()
        # 重新连接下拉框信号
        self.pier_combo.currentIndexChanged.connect(self.switch_pier)
        # 更新汇总信息表
        self.update_summary()
    def reorder_components(self, deleted_number):
        """重新排序所有构件编号
        Args:
            deleted_number: 被删除的构件编号
        """
        # 分别获取桥台和桥墩的信息
        abutments = [(comp['number'], comp) for comp in self.all_components
                    if comp['type'] == "桥台"]
        piers = [(comp['number'], comp) for comp in self.all_components
                if comp['type'] == "桥墩"]
        # 分别排序
        abutments.sort()
        piers.sort()
        # 创建新的数据存储
        new_abutment_data = {}
        new_pier_data = {}
        # 处理桥墩编号调整
        # 移除了 enumerate，因为索引 i 未被使用
        for old_num, comp in piers:
            # 只调整大于被删除编号的构件
            if old_num > deleted_number and comp['type'] == "桥墩":
                # 新编号 = 旧编号 - 1
                new_num = old_num - 1
                comp['number'] = new_num
                # 保留原数据
                if old_num in self.pier_data:
                    new_pier_data[new_num] = self.pier_data[old_num]
            else:
                # 不需要调整的构件保持原编号
                new_pier_data[old_num] = self.pier_data.get(old_num, {})
        # 处理桥台编号调整
        # 移除了 enumerate，因为索引 i 未被使用
        for old_num, comp in abutments:
            # 只调整大于被删除编号的构件
            if old_num > deleted_number and comp['type'] == "桥台":
                # 新编号 = 旧编号 - 1
                new_num = old_num - 1
                comp['number'] = new_num
                # 保留原数据
                if old_num in self.abutment_data:
                    new_abutment_data[new_num] = self.abutment_data[old_num]
            else:
                # 不需要调整的构件保持原编号
                new_abutment_data[old_num] = self.abutment_data.get(old_num, {})
        # 更新数据存储
        self.abutment_data = new_abutment_data
        self.pier_data = new_pier_data
    def switch_pier(self, _index):
        """切换桥台/桥墩编号"""
        # 保存当前数据（确保在切换前保存）
        if self.current_pier is not None:
            self.save_current_data()
        text = self.pier_combo.currentText()
        if text == "请添加":
            self.current_pier = None
            return
        # 获取当前编号
        if text.startswith("第") and text.endswith(("台", "墩")):
            try:
                self.current_pier = int(text[1:-2])
                # 加载新数据
                self.load_structure_data()
                # 更新汇总信息
                self.update_summary()
            except ValueError:
                QMessageBox.warning(self, "编号解析错误", f"无法解析编号: {text}")
                self.current_pier = None
    def clear_table_data(self):
        """清空表格数据，保留结构"""
        structure_config = self.config["下部结构"][self.structure_type]
        for row in range(self.table.rowCount()):
            if self.table.item(row, 1):  # 检查损坏类型单元格是否存在
                damage_type = self.table.item(row, 1).text()
                # 在配置中查找该损坏类型的第一个选项
                for part_data in structure_config.values():
                    if damage_type in part_data:
                        default_option = part_data[damage_type]["options"][0]
                        break
                else:
                    default_option = "无"
                # 重置损坏程度下拉框为默认选项
                if self.table.cellWidget(row, 2):
                    combo = self.table.cellWidget(row, 2)
                    combo.setCurrentText(default_option)
            # 重置数值单元格
            for col in [3, 4, 5, 6, 7]:
                if not self.table.item(row, col):
                    item = QTableWidgetItem('0' if col == 3 else '0.00')
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(row, col, item)
                else:
                    self.table.item(row, col).setText('0' if col == 3 else '0.00')
        # 重置BCI值
        if not self.table.item(0, 8):
            item = QTableWidgetItem('100.00')
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(0, 8, item)
        else:
            self.table.item(0, 8).setText('100.00')
        # 重新计算构件权重
        self.calculate_component_weights()
    def update_summary(self):
        """更新汇总信息"""
        try:
            # 保存当前数据，确保最新状态被记录
            if self.current_pier is not None:
                self.save_current_data()
            # 清空汇总表格
            self.summary_table.setRowCount(0)
            # 获取所有有效的桥台和桥墩数据
            all_components = []
            # 添加桥台数据（直接从abutment_data获取，不依赖all_components）
            for number, data in self.abutment_data.items():
                if isinstance(number, (int, float)):  # 确保编号是数字
                    # 确保BCI值是最新的
                    bci_value = data.get('bci', 100.00)
                    # 如果当前正在编辑这个构件，从表格获取最新值
                    if self.structure_type == '桥台' and self.current_pier == number and self.table.item(0, 8):
                        try:
                            current_bci = float(self.table.item(0, 8).text())
                            if current_bci != bci_value:
                                # 更新数据结构中的值
                                self.abutment_data[number]['bci'] = current_bci
                                bci_value = current_bci
                        except (ValueError, TypeError):
                            pass
                    all_components.append({
                        'number': number,
                        'type': '桥台',
                        'bci': bci_value
                    })
            # 添加桥墩数据（直接从pier_data获取，不依赖all_components）
            for number, data in self.pier_data.items():
                if isinstance(number, (int, float)):  # 确保编号是数字
                    # 确保BCI值是最新的
                    bci_value = data.get('bci', 100.00)
                    # 如果当前正在编辑这个构件，从表格获取最新值
                    if self.structure_type == '桥墩' and self.current_pier == number and self.table.item(0, 8):
                        try:
                            current_bci = float(self.table.item(0, 8).text())
                            if current_bci != bci_value:
                                # 更新数据结构中的值
                                self.pier_data[number]['bci'] = current_bci
                                bci_value = current_bci
                        except (ValueError, TypeError):
                            pass
                    all_components.append({
                        'number': number,
                        'type': '桥墩',
                        'bci': bci_value
                    })
            if not all_components:
                return
            # 按编号排序
            all_components.sort(key=lambda x: x['number'])
            # 计算BCI(xj)的平均值和最小值
            bci_values = [comp['bci'] for comp in all_components]
            bci_average = sum(bci_values) / len(bci_values) if bci_values else 100.00
            bci_min = min(bci_values) if bci_values else 100.00
            # 添加所有构件数据到汇总表
            for comp in all_components:
                row = self.summary_table.rowCount()
                self.summary_table.insertRow(row)
                # 添加构件编号和类型
                pier_item = QTableWidgetItem(f"第{comp['number']}#{comp['type'][-1]}")
                pier_item.setTextAlignment(Qt.AlignCenter)
                self.summary_table.setItem(row, 0, pier_item)
                # 添加BCI(xj)值
                bci_item = QTableWidgetItem(f"{comp['bci']:.2f}")
                bci_item.setTextAlignment(Qt.AlignCenter)
                self.summary_table.setItem(row, 1, bci_item)
                # 添加BCI(x)值和BSI(x)值（仅在第一行）
                if row == 0:
                    bci_x_item = QTableWidgetItem(f"{bci_average:.2f}")
                    bci_x_item.setTextAlignment(Qt.AlignCenter)
                    self.summary_table.setItem(row, 2, bci_x_item)
                    bsi_item = QTableWidgetItem(f"{bci_min:.2f}")
                    bsi_item.setTextAlignment(Qt.AlignCenter)
                    self.summary_table.setItem(row, 3, bsi_item)
            # 合并单元格
            if len(all_components) > 1:
                self.summary_table.setSpan(0, 2, len(all_components), 1)
                self.summary_table.setSpan(0, 3, len(all_components), 1)
        except Exception as e:
            QMessageBox.warning(self, "更新汇总错误", f"更新汇总信息时出错: {str(e)}")
    def init_table(self):
        """初始化表格"""
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        headers = ['部位', '损坏类型', '损坏程度', '扣分值DP(jkl)', '权重ω(jkl)', 'DP(jkl)*ω(jkl)', 'SDP(jk)',
                   '权重ω(jk)', 'BCI(xj)']
        self.table.setHorizontalHeaderLabels(headers)
        # 设置表格基本属性
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.table.horizontalHeader().setFixedHeight(25)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.verticalHeader().setVisible(False)
        # 设置列宽
        column_widths = [107, 127, 127, 107, 107, 107, 75, 75, 76]
        for col, width in enumerate(column_widths):
            self.table.setColumnWidth(col, width)
            self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Fixed)
        # 设置表格样式
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 4px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 2px 4px;
            }
        """)
    def switch_structure_type(self):
        """切换结构类型（桥台/桥墩）"""
        # print("\n===== 开始切换结构类型 =====")
        # print(f"当前结构类型: {self.structure_type}, 当前构件编号: {self.current_pier}")
        
        # 保存当前数据
        if self.current_pier is not None:
            # print(f"保存当前数据: {self.structure_type} #{self.current_pier}")
            self.save_current_data()
        
        # 更新结构类型
        old_structure_type = self.structure_type
        self.structure_type = "桥台" if self.abutment_radio.isChecked() else "桥墩"
        # print(f"切换结构类型: {old_structure_type} -> {self.structure_type}")
        
        # 如果结构类型没有变化，直接返回
        if old_structure_type == self.structure_type:
            return
        
        # 安全地断开所有下拉框的信号连接
        for row in range(self.table.rowCount()):
            if self.table.cellWidget(row, 2):
                combo = self.table.cellWidget(row, 2)
                try:
                    signals = combo.receivers(combo.currentTextChanged)
                    if signals > 0:
                        combo.currentTextChanged.disconnect()
                except Exception:
                    pass
        
        # 重置当前构件编号（但不清除已保存的数据）
        self.current_pier = None
        
        # 更新下拉框
        self.update_pier_combo()
        
        # 如果当前类型有已保存的数据，加载第一个构件的数据
        current_data = self.abutment_data if self.structure_type == "桥台" else self.pier_data
        if current_data:
            valid_keys = [k for k in current_data.keys() if k is not None and isinstance(k, (int, float))]
            if valid_keys:
                first_pier = min(valid_keys)
                # print(f"找到 {self.structure_type} 的有效构件: {valid_keys}")
                # print(f"选择第一个构件: #{first_pier}")
                self.current_pier = first_pier
                
                # 打印当前数据内容
                if first_pier in current_data:
                    data = current_data[first_pier]
                    # print(f"构件 #{first_pier} 的数据行数: {len(data.get('table_data', []))}")
                    for i, row_data in enumerate(data.get('table_data', [])):
                        damage_type = row_data.get('damage_type')
                        damage_degree = row_data.get('damage_degree', '无')
                        # print(f"  行 {i}: 损坏类型={damage_type}, 程度={damage_degree}")
                
                # 先初始化表格结构，但不清空数据
                # print("初始化表格结构")
                self.init_sub_data(skip_clear=True)
                
                # 直接加载保存的数据
                # print("加载保存的数据")
                self.load_structure_data()
        else:
            # print(f"未找到 {self.structure_type} 的已保存数据")
            # 如果没有数据，初始化表格结构
            self.init_sub_data()
        
        # print("===== 结构类型切换完成 =====\n")
    def save_current_data(self):
        """保存当前数据到对应的数据结构"""
        current_data = {
            'table_data': [],
            'bci': float(self.table.item(0, 8).text()) if self.table.item(0, 8) else 100.00,
            'checkbox_states': {}
        }
        
        # print(f"\n===== 开始保存 {self.structure_type} #{self.current_pier} 数据 =====")
        
        for row in range(self.table.rowCount()):
            # 保存复选框状态
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    current_data['checkbox_states'][row] = checkbox.isChecked()
            
            # 保存每一行的数据（不管是否为空）
            if self.table.item(row, 1):  # 只要损坏类型单元格存在
                damage_type = self.table.item(row, 1).text()
                combo = self.table.cellWidget(row, 2)
                damage_degree = combo.currentText() if combo and isinstance(combo, QComboBox) else "无"
                
                # 打印当前行的损坏类型和程度
                # print(f"  行 {row}: 损坏类型={damage_type}, 程度={damage_degree}")
                
                # 创建行数据
                row_data = {
                    'damage_type': damage_type,
                    'damage_degree': damage_degree,
                    'deduct_points': self.table.item(row, 3).text() if self.table.item(row, 3) else "0",
                    'weight': self.table.item(row, 4).text() if self.table.item(row, 4) else "0.00",
                    'product': self.table.item(row, 5).text() if self.table.item(row, 5) else "0.00",
                    'sdp': self.table.item(row, 6).text() if self.table.item(row, 6) else "0.00",
                    'component_weight': self.table.item(row, 7).text() if self.table.item(row, 7) else "0.00"
                }
                current_data['table_data'].append(row_data)
        
        # 根据当前类型保存到对应的数据结构
        if self.structure_type == "桥台":
            self.abutment_data[self.current_pier] = current_data
            # print(f"  保存到 abutment_data[{self.current_pier}], 共 {len(current_data['table_data'])} 行数据")
        else:
            self.pier_data[self.current_pier] = current_data
        #     print(f"  保存到 pier_data[{self.current_pier}], 共 {len(current_data['table_data'])} 行数据")
        #
        # print("===== 数据保存完成 =====\n")
    def load_structure_data(self):
        """加载对应类型的数据"""
        print(f"\n===== 开始加载 {self.structure_type} #{self.current_pier} 的数据 =====")
        
        # 获取当前类型的数据
        current_data = (self.abutment_data if self.structure_type == "桥台"
                        else self.pier_data).get(self.current_pier)
        
        if not current_data:
            print(f"警告：未找到 {self.structure_type} #{self.current_pier} 的数据")
            self.clear_table_data()
            return
        
        # 打印保存的数据内容
        print(f"找到数据行数: {len(current_data.get('table_data', []))}")
        
        # 创建损坏类型映射表（用于桥墩和桥台之间的转换）
        damage_type_mapping = {
            # 桥墩 -> 桥台
            "盖梁混凝土剥离": "台帽混凝土剥离",
            "墩身水平裂缝": "台身水平裂缝",
            "墩身纵向裂缝": "台身纵向裂缝",
            "墩身混凝土剥离": "台身混凝土剥离",
            "桥墩倾斜": "桥台倾斜",
            # 桥台 -> 桥墩
            "台帽混凝土剥离": "盖梁混凝土剥离",
            "台身水平裂缝": "墩身水平裂缝",
            "台身纵向裂缝": "墩身纵向裂缝",
            "台身混凝土剥离": "墩身混凝土剥离",
            "桥台倾斜": "桥墩倾斜"
        }
        
        # 定义耳背翼墙构件的所有损坏类型
        ear_wing_damage_types = [
            "剥离脱落",
            "翼墙前结合处开裂",
            "挡土功能缺失",
            "翼墙大贯通缝"
        ]
        
        # 定义耳背翼墙构件各损坏类型的正常状态
        ear_wing_normal_states = {
            "剥离脱落": "无",
            "翼墙前结合处开裂": "完好",
            "挡土功能缺失": "完好",
            "翼墙大贯通缝": "无"
        }
        
        # 创建损坏类型到行号的映射
        damage_type_to_row = {}
        for row in range(self.table.rowCount()):
            if self.table.item(row, 1):
                damage_type = self.table.item(row, 1).text()
                damage_type_to_row[damage_type] = row
                print(f"  表格行 {row}: 损坏类型={damage_type}")
        # 断开所有下拉框的信号连接
        for row in range(self.table.rowCount()):
            if self.table.cellWidget(row, 2):
                combo = self.table.cellWidget(row, 2)
                try:
                    combo.currentTextChanged.disconnect()
                except Exception:
                    pass
        

        # 如果是桥台，确保耳背翼墙的损坏类型已添加到表格中
        self.current_row = 20  # 初始值可以根据实际情况调整
        if self.structure_type == "桥台":
            # 获取当前桥台的复选框状态数据
            current_checkbox_states = current_data.get('checkbox_states', {}) if current_data else {}
            # 确保表格有足够的行数
            if self.table.rowCount() < self.current_row + len(ear_wing_damage_types):
                self.table.setRowCount(self.current_row + len(ear_wing_damage_types))
            
            for damage_type in ear_wing_damage_types:
                if damage_type not in damage_type_to_row:
                    start_row = self.current_row
                    # 获取已存在的复选框单元格
                    cell_widget = self.table.cellWidget(start_row, 0)
                    if cell_widget:
                        checkbox = cell_widget.findChild(QCheckBox)
                    if checkbox:
                        # 根据保存的状态设置复选框
                        is_checked = current_checkbox_states.get(start_row, True)  # 如果没有保存的状态，默认选中
                        checkbox.setChecked(is_checked)
                        print(f"  设置耳背翼墙损坏类型复选框状态: {damage_type} -> 行 {row}, 状态: {is_checked}")
                    # 在对应行号添加损坏类型
                    damage_item = QTableWidgetItem(damage_type)
                    damage_item.setTextAlignment(Qt.AlignCenter)
                    damage_item.setFlags(damage_item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(start_row, 1, damage_item)
                    damage_type_to_row[damage_type] = start_row
                    print(f"  添加耳背翼墙损坏类型到表格: {damage_type} -> 行 {start_row}")
                    
                    # 为新添加的行创建下拉框
                    combo = QComboBox()
                    # 根据损坏类型设置下拉框选项
                    if damage_type == "剥离脱落":
                        options = ["无", "轻微", "严重"]
                    elif damage_type == "翼墙前结合处开裂":
                        options = ["完好", "开裂", "脱开"]
                    elif damage_type == "挡土功能缺失":
                        options = ["完好", "失去部分", "完全散失"]
                    elif damage_type == "翼墙大贯通缝":
                        options = ["无", "少量", "大量"]
                    
                    combo.addItems(options)
                    combo.setEditable(True)
                    combo.lineEdit().setAlignment(Qt.AlignCenter)
                    combo.lineEdit().setReadOnly(True)
                    
                    # 设置下拉框样式
                    combo.setStyleSheet("""
                        QComboBox {
                            border: 1px solid #ddd;
                            border-radius: 3px;
                            padding: 1px 4px 1px 3px;
                            background: white;
                            min-width: 100px;
                        }
                        QComboBox::drop-down {
                            border: none;
                            border-left: 1px solid #ddd;
                            width: 20px;
                        }
                        QComboBox::down-arrow {
                            width: 8px;
                            height: 8px;
                            margin-top: 1px;
                            margin-right: 4px;
                            border-top: 2px solid #666;
                            border-right: 2px solid #666;
                            transform: rotate(135deg);
                        }
                        QComboBox QAbstractItemView {
                            border: 1px solid #ddd;
                            selection-background-color: #f0f0f0;
                            min-width: 70px;
                        }
                        QComboBox QAbstractItemView::item {
                            height: 24px;
                            padding: 4px;
                            text-align: center;
                        }
                    """)
                    
                    # 初始化所有单元格，确保它们存在
                    for col in range(3, 8):
                        default_value = '0' if col == 3 else '0.00'
                        item_widget = QTableWidgetItem(str(default_value))
                        item_widget.setTextAlignment(Qt.AlignCenter)
                        item_widget.setFlags(item_widget.flags() & ~Qt.ItemIsEditable)
                        self.table.setItem(start_row, col, item_widget)
                    
                    # 设置下拉框到表格中
                    self.table.setCellWidget(start_row, 2, combo)
                    
                    # 强制处理事件，确保组件被正确添加
                    QApplication.processEvents()
                    
                    # 设置信号连接
                    combo.currentTextChanged.connect(
                        lambda text, r=start_row: self.update_deduct_points(
                            self.table.item(r, 1).text(), r
                        )
                    )
                    
                    # 验证下拉框创建
                    widget = self.table.cellWidget(start_row, 2)
                    print(f"  验证下拉框创建: 行 {start_row}, 组件类型: {type(widget)}")
                    
                    # 递增当前行号
                    self.current_row += 1
        
        # 初始化所有单元格和BCI值
        self.init_table_cells()
        
        # 在init_table_cells后重新检查并创建耳背翼墙的下拉框
        if self.structure_type == "桥台":
            for damage_type in ear_wing_damage_types:
                if damage_type in damage_type_to_row:
                    row = damage_type_to_row[damage_type]
                    # 如果下拉框不存在或类型不正确，重新创建
                    if not isinstance(self.table.cellWidget(row, 2), QComboBox):
                        # 为行创建新的下拉框
                        combo = QComboBox()
                        # 根据损坏类型设置下拉框选项
                        if damage_type == "剥离脱落":
                            options = ["无", "轻微", "严重"]
                        elif damage_type == "翼墙前结合处开裂":
                            options = ["完好", "开裂", "脱开"]
                        elif damage_type == "挡土功能缺失":
                            options = ["完好", "失去部分", "完全散失"]
                        elif damage_type == "翼墙大贯通缝":
                            options = ["无", "少量", "大量"]
                        
                        combo.addItems(options)
                        combo.setEditable(True)
                        combo.lineEdit().setAlignment(Qt.AlignCenter)
                        combo.lineEdit().setReadOnly(True)
                        
                        # 设置下拉框样式
                        combo.setStyleSheet("""
                            QComboBox {
                                border: 1px solid #ddd;
                                border-radius: 3px;
                                padding: 1px 4px 1px 3px;
                                background: white;
                                min-width: 100px;
                            }
                            QComboBox::drop-down {
                                border: none;
                                border-left: 1px solid #ddd;
                                width: 20px;
                            }
                            QComboBox::down-arrow {
                                width: 8px;
                                height: 8px;
                                margin-top: 1px;
                                margin-right: 4px;
                                border-top: 2px solid #666;
                                border-right: 2px solid #666;
                                transform: rotate(135deg);
                            }
                            QComboBox QAbstractItemView {
                                border: 1px solid #ddd;
                                selection-background-color: #f0f0f0;
                                min-width: 70px;
                            }
                            QComboBox QAbstractItemView::item {
                                height: 24px;
                                padding: 4px;
                                text-align: center;
                            }
                        """)
                        
                        # 设置下拉框到表格中
                        self.table.setCellWidget(row, 2, combo)
                        
                        # 强制处理事件
                        QApplication.processEvents()
                        
                        # 设置信号连接
                        combo.currentTextChanged.connect(
                            lambda text, r=row: self.update_deduct_points(
                                self.table.item(r, 1).text(), r
                            )
                        )
                        
                        print(f"  重新创建下拉框: 行 {row}, 损坏类型 {damage_type}, 组件: {type(self.table.cellWidget(row, 2))}")
        # 检查是否所有保存的损坏类型都能在表格中找到
        print("\n检查损坏类型映射：")
        for i, row_data in enumerate(current_data.get('table_data', [])):
            damage_type = row_data.get('damage_type')
            damage_degree = row_data.get('damage_degree', '无')
            print(f"  保存的数据 {i}: 损坏类型={damage_type}, 程度={damage_degree}")
            
            # 尝试直接匹配
            if damage_type in damage_type_to_row:
                row = damage_type_to_row[damage_type]
                print(f"    找到映射: {damage_type} -> 行 {row}")
            # 尝试通过映射表匹配
            elif damage_type in damage_type_mapping and damage_type_mapping[damage_type] in damage_type_to_row:
                mapped_type = damage_type_mapping[damage_type]
                row = damage_type_to_row[mapped_type]
                print(f"    通过映射找到: {damage_type} -> {mapped_type} -> 行 {row}")
                # 更新损坏类型，使其与当前表格匹配
                row_data['damage_type'] = mapped_type
            else:
                print(f"    警告: 找不到损坏类型 '{damage_type}' 的映射")
        
        # 先恢复复选框状态，因为这会影响权重计算
        for row, is_checked in current_data.get('checkbox_states', {}).items():
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    checkbox.setChecked(is_checked)
        # 遍历保存的数据，恢复到对应的行
        print("\n开始恢复数据：")
        for row_data in current_data.get('table_data', []):
            damage_type = row_data.get('damage_type')
            if damage_type in damage_type_to_row:
                row = damage_type_to_row[damage_type]
                # 设置损坏程度（下拉框）
                combo = self.table.cellWidget(row, 2)
                if combo and isinstance(combo, QComboBox):
                    damage_degree = row_data.get('damage_degree', '无')
                    print(f"  尝试设置 '{damage_type}' (行 {row}) 的损坏程度为 '{damage_degree}'")
                    
                    # 检查下拉框中是否有该选项
                    index = combo.findText(damage_degree)
                    if index >= 0:
                        print(f"    找到选项索引: {index}")
                        combo.setCurrentIndex(index)
                    else:
                        print(f"    警告: 下拉框中找不到选项 '{damage_degree}'")
                        # 尝试直接设置文本
                        combo.setCurrentText(damage_degree)
                else:
                    print(f"  警告: 行 {row} 没有有效的下拉框")
                
                # 设置其他数值列
                for col, key in enumerate(['deduct_points', 'weight', 'product', 'sdp', 'component_weight'], 3):
                    if self.table.item(row, col):
                        self.table.item(row, col).setText(str(row_data.get(key, '0')))   
        # 重新连接所有下拉框的信号
        self.reconnect_combo_signals()
        # 最后更新BCI值，避免中间状态的重复计算
        if 'bci' in current_data and self.table.item(0, 8):
            self.table.item(0, 8).setText(f"{current_data['bci']:.2f}")
        # 强制重新计算构件权重，这会触发BCI值的更新
        self.calculate_component_weights()
        print("===== 数据加载完成 =====\n")

    def init_table_cells(self):
        """初始化表格单元格"""
        for row in range(self.table.rowCount()):
            # 初始化数值单元格
            for col in [3, 4, 5, 6, 7]:
                if not self.table.item(row, col):
                    item = QTableWidgetItem('0' if col == 3 else '0.00')
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(row, col, item)
        # 确保BCI单元格存在
        if not self.table.item(0, 8):
            item = QTableWidgetItem('100.00')
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(0, 8, item)
            self.table.setSpan(0, 8, self.table.rowCount(), 1)
    def restore_checkbox_states(self, current_data):
        """恢复复选框状态"""
        for row, is_checked in current_data.get('checkbox_states', {}).items():
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    checkbox.setChecked(is_checked)
    def restore_table_data(self, current_data):
        """恢复表格数据"""
        # 获取当前结构类型的配置
        structure_config = self.config["下部结构"][self.structure_type]
        # 创建损坏类型到行号的映射
        damage_type_to_row = {}
        for row in range(self.table.rowCount()):
            if self.table.item(row, 1):
                damage_type = self.table.item(row, 1).text()
                damage_type_to_row[damage_type] = row
        # 遍历保存的数据，恢复到对应的行
        for row_data in current_data.get('table_data', []):
            damage_type = row_data.get('damage_type')
            if damage_type in damage_type_to_row:
                row = damage_type_to_row[damage_type]
                # 设置损坏程度
                if self.table.cellWidget(row, 2):
                    combo = self.table.cellWidget(row, 2)
                    # 确保下拉框有正确的选项
                    for part_data in structure_config.values():
                        if damage_type in part_data:
                            options = part_data[damage_type]["options"]
                            # 重新创建下拉框选项
                            combo.clear()
                            combo.addItems(options)
                            break
                    # 设置保存的选择
                    damage_degree = row_data.get('damage_degree', "无")
                    combo.setCurrentText(damage_degree)
                # 恢复其他数值
                columns = {
                    3: ('deduct_points', '0'),
                    4: ('weight', '0.00'),
                    5: ('product', '0.00'),
                    6: ('sdp', '0.00'),
                    7: ('component_weight', '0.00')
                }
                for col, (key, default_value) in columns.items():
                    if self.table.item(row, col):
                        self.table.item(row, col).setText(row_data.get(key, default_value))
    def reconnect_combo_signals(self):
        """重新连接下拉框信号"""
        for row in range(self.table.rowCount()):
            if self.table.cellWidget(row, 2):
                combo = self.table.cellWidget(row, 2)
                combo.currentTextChanged.connect(
                    lambda text, r=row: self.update_deduct_points(
                        self.table.item(r, 1).text(), r
                    )
                )
    def update_pier_combo(self):
        """更新下拉框选项"""
        self.pier_combo.clear()
        # 获取当前类型的构件
        current_components = [comp for comp in self.all_components
                              if comp['type'] == self.structure_type]
        if not current_components:
            # 只有在没有构件时才显示"请添加"
            self.pier_combo.addItem("请添加")
            self.current_pier = None
            # 禁用删除按钮
            self.del_btn.setEnabled(False)
        else:
            # 添加当前类型的已有选项
            for comp in sorted(current_components, key=lambda x: x['number']):
                self.pier_combo.addItem(
                    f"第{comp['number']}#{'台' if self.structure_type == '桥台' else '墩'}")
            # 启用删除按钮（因为存在可删除的构件）
            self.del_btn.setEnabled(True)
            # 设置第一个构件为当前选中项
            self.pier_combo.setCurrentIndex(0)
    def load_sub_config(self):
        """加载下部结构配置"""
        sub_structure = self.config["下部结构"]
        for structure_type in ["桥台", "桥墩"]:
            for part, damages in sub_structure[structure_type].items():
                for damage_type, info in damages.items():
                    self.deduct_points[damage_type] = info["scores"]
    def get_part_range(self, damage_type):
        """获取下部结构部位范围"""
        if self.structure_type == "桥台":
            # 桥台的损坏类型对应范围
            ranges = {
                "台帽": (["表面裂缝", "台帽混凝土剥离", "露筋锈蚀", "结构裂缝", "裂缝处渗水", "墩台成块剥落"], (0, 6)),
                "台身": (
                    ["台身水平裂缝", "台身纵向裂缝", "框架式节点裂缝", "墩台露筋锈蚀", "台身混凝土剥离", "桥台倾斜",
                     "桥面贯通横缝"], (6, 13)),
                "基础": (["基础冲刷", "基础掏空", "混凝土桩损坏", "基础位移"], (13, 17)),
                "拱脚": (["拱脚位移", "拱脚开裂", "拱脚残缺"], (17, 20)),
                "耳背翼墙": (["剥离脱落", "翼墙前结合处开裂", "挡土功能缺失", "翼墙大贯通缝"], (20, 24))
            }
        else:  # 桥墩
            ranges = {
                "盖梁": (["表面裂缝", "盖梁混凝土剥离", "露筋锈蚀", "结构裂缝", "裂缝处渗水", "墩台成块剥落"], (0, 6)),
                "墩身": (
                    ["墩身水平裂缝", "墩身纵向裂缝", "框架式节点裂缝", "墩台露筋锈蚀", "墩身混凝土剥离", "桥墩倾斜",
                     "桥面贯通横缝"], (6, 13)),
                "基础": (["基础冲刷", "基础掏空", "混凝土桩损坏", "基础位移"], (13, 17)),
                "拱脚": (["拱脚位移", "拱脚开裂", "拱脚残缺"], (17, 20)),
            }
        # 遍历查找损坏类型所属范围
        for part_name, (damage_types, range_tuple) in ranges.items():
            if damage_type in damage_types:
                return range_tuple
        # 如果没有找到匹配的范围，抛出异常
        raise ValueError(f"无法找到损坏类型 '{damage_type}' 的对应范围。当前结构类型：{self.structure_type}")
    def calculate_component_weights(self):
        """计算各构件权重"""
        # 桥台和桥墩的标准权重
        standard_weights = {
            "桥台": {
                "台帽": 0.10,
                "台身": 0.30,
                "基础": 0.35,
                "拱脚": 0.15,
                "耳背翼墙": 0.10
            },
            "桥墩": {
                "盖梁": 0.10,
                "墩身": 0.30,
                "基础": 0.45,
                "拱脚": 0.15
            }
        }
        current_weights = standard_weights[self.structure_type]
        # 检查每个构件是否存在
        existing_components = {}
        total_standard_weight = 0
        row = 0
        while row < self.table.rowCount():
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    component_name = checkbox.text()
                    if not checkbox.isChecked():
                        self.table.item(row, 7).setText("0.00")
                    elif component_name in current_weights:
                        existing_components[component_name] = current_weights[component_name]
                        total_standard_weight += current_weights[component_name]
            row += self.table.rowSpan(row, 0)
        if not existing_components:
            return
        # 重新分配权重
        weights = []
        for component, weight in existing_components.items():
            adjusted_weight = weight / total_standard_weight
            weights.append((component, adjusted_weight))
        # 处理舍入误差
        total_rounded = sum(round(w[1], 2) for w in weights)
        if total_rounded != 1:
            max_component = max(weights, key=lambda x: x[1])
            weights = [(c, round(w, 2) if c != max_component[0] else
            round(w + (1 - total_rounded), 2)) for c, w in weights]
        # 更新表格中的权重值
        row = 0
        while row < self.table.rowCount():
            cell_widget = self.table.cellWidget(row, 0)
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    component_name = checkbox.text()
                    for name, weight in weights:
                        if name == component_name:
                            self.table.item(row, 7).setText(f"{weight:.2f}")
                            break
            row += self.table.rowSpan(row, 0)
        # 更新BCI值
        self.calculate_bsi_bci()
    def calculate_bsi_bci(self):
        """重写计算方法（只计算BCI）"""
        mdp_values = []
        weighted_sum = 0
        row = 0
        while row < self.table.rowCount():
            if self.table.item(row, 6):  # 检查SDP(ij)列
                mdp = float(self.table.item(row, 6).text())
                weight = float(self.table.item(row, 7).text())
                mdp_values.append(mdp)
                weighted_sum += mdp * weight
            row += self.table.rowSpan(row, 0)
        # 计算BCI值
        bci = 100 - weighted_sum
        self.table.item(0, 8).setText(f"{bci:.2f}")
        # 保存当前数据并更新汇总信息
        if self.current_pier is not None:
            self.save_current_data()
            self.update_summary()
    def init_sub_data(self, skip_clear=False):
        """初始化下部结构数据"""
        # 如果需要保留数据，先保存当前表格状态
        saved_data = None
        if skip_clear and self.current_pier is not None:
            # 直接使用已保存的数据，而不是从表格中提取
            if self.structure_type == "桥台":
                saved_data = self.abutment_data.get(self.current_pier)
            else:
                saved_data = self.pier_data.get(self.current_pier)
                
            # 打印保存的数据，用于调试
            if saved_data:
                print(f"\n===== 保存的 {self.structure_type} #{self.current_pier} 数据 =====")
                for i, row_data in enumerate(saved_data.get('table_data', [])):
                    damage_type = row_data.get('damage_type')
                    damage_degree = row_data.get('damage_degree', '无')
                    print(f"  行 {i}: 损坏类型={damage_type}, 程度={damage_degree}")
        
        # 只有在不需要保留数据时才清空表格
        if not skip_clear:
            self.table.setRowCount(0)
        
        # 从配置文件构建数据结构，确保使用正确的结构类型
        data = []
        structure_config = self.config["下部结构"][self.structure_type]
        for part_name, part_data in structure_config.items():
            data.append({
                'name': part_name,
                'items': list(part_data.keys())
            })
        
        # 计算总行数
        total_rows = sum(len(group['items']) for group in data)
        self.table.setRowCount(total_rows)
        
        # 创建损坏类型到保存数据的映射
        saved_damage_data = {}
        if skip_clear and saved_data:
            for row_data in saved_data.get('table_data', []):
                damage_type = row_data.get('damage_type')
                if damage_type:
                    saved_damage_data[damage_type] = row_data
        
        # 填充数据并合并单元格
        current_row = 0
        for group in data:
            start_row = current_row
            # 填充损坏类型
            for item in group['items']:
                # 创建损坏类型单元格
                damage_item = QTableWidgetItem(item)
                damage_item.setTextAlignment(Qt.AlignCenter)
                damage_item.setFlags(damage_item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(current_row, 1, damage_item)
                
                # 创建下拉框
                combo = QComboBox()
                options = structure_config[group['name']][item]["options"]
                combo.addItems(options)
                combo.setEditable(True)
                combo.lineEdit().setAlignment(Qt.AlignCenter)
                combo.lineEdit().setReadOnly(True)
                
                # 如果有保存的数据，设置下拉框的值
                if skip_clear and item in saved_damage_data:
                    saved_degree = saved_damage_data[item].get('damage_degree', '无')
                    index = combo.findText(saved_degree)
                    if index >= 0:
                        combo.setCurrentIndex(index)
                        print(f"  设置下拉框 '{item}' 的值为 '{saved_degree}' (索引 {index})")
                
                # 设置下拉框的信号连接
                current_row_copy = current_row
                combo.currentTextChanged.connect(
                    lambda text, r=current_row_copy: self.update_deduct_points(
                        self.table.item(r, 1).text(), r
                    )
                )
                
                # 设置下拉框样式
                combo.setStyleSheet("""
                    QComboBox {
                        border: 1px solid #ddd;
                        border-radius: 3px;
                        padding: 1px 4px 1px 3px;
                        background: white;
                        min-width: 100px;
                    }
                    QComboBox::drop-down {
                        border: none;
                        border-left: 1px solid #ddd;
                        width: 20px;
                    }
                    QComboBox::down-arrow {
                        width: 8px;
                        height: 8px;
                        margin-top: 1px;
                        margin-right: 4px;
                        border-top: 2px solid #666;
                        border-right: 2px solid #666;
                        transform: rotate(135deg);
                    }
                    QComboBox QAbstractItemView {
                        border: 1px solid #ddd;
                        selection-background-color: #f0f0f0;
                        min-width: 70px;
                    }
                    QComboBox QAbstractItemView::item {
                        height: 24px;
                        padding: 4px;
                        text-align: center;
                    }
                """)
                self.table.setCellWidget(current_row, 2, combo)
                
                # 初始化数值单元格
                for col in range(3, 8):
                    default_value = '0' if col == 3 else '0.00'
                    
                    # 如果有保存的数据，使用保存的值
                    if skip_clear and item in saved_damage_data:
                        if col == 3:
                            default_value = saved_damage_data[item].get('deduct_points', '0')
                        elif col == 4:
                            default_value = saved_damage_data[item].get('weight', '0.00')
                        elif col == 5:
                            default_value = saved_damage_data[item].get('product', '0.00')
                        elif col == 6:
                            default_value = saved_damage_data[item].get('sdp', '0.00')
                        elif col == 7:
                            default_value = saved_damage_data[item].get('component_weight', '0.00')
                    
                    item_widget = QTableWidgetItem(str(default_value))
                    item_widget.setTextAlignment(Qt.AlignCenter)
                    item_widget.setFlags(item_widget.flags() & ~Qt.ItemIsEditable)
                    self.table.setItem(current_row, col, item_widget)
                
                current_row += 1
            
            # 创建部位复选框
            self.table.setCellWidget(start_row, 0, self.create_component_cell(group['name']))
            
            # 设置合并单元格
            if len(group['items']) > 1:
                # 合并第0列（部位）
                self.table.setSpan(start_row, 0, len(group['items']), 1)
                # 合并第6列（SDP(jk)）
                self.table.setSpan(start_row, 6, len(group['items']), 1)
                # 合并第7列（权重ω(jk)）
                self.table.setSpan(start_row, 7, len(group['items']), 1)
        
        # 设置BCI单元格
        item = QTableWidgetItem('100.00')
        item.setTextAlignment(Qt.AlignCenter)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(0, 8, item)
        self.table.setSpan(0, 8, total_rows, 1)
        
        # 如果有保存的数据，恢复复选框状态和BCI值
        if skip_clear and saved_data:
            # 恢复复选框状态
            self.restore_checkbox_states(saved_data)
            
            # 恢复BCI值
            if 'bci' in saved_data and self.table.item(0, 8):
                self.table.item(0, 8).setText(f"{saved_data['bci']:.2f}")
            
            # 不需要调用restore_table_data，因为我们已经在创建单元格时设置了值
            print("  跳过restore_table_data调用，已直接设置单元格值")
        else:
            # 重新计算权重
            self.calculate_component_weights()
    def create_component_cell(self, name):
        """创建带复选框的部位单元格"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(5, 2, 5, 2)
        checkbox = QCheckBox(name)
        checkbox.setChecked(True)  # 默认选中
        checkbox.stateChanged.connect(lambda: self.calculate_component_weights())
        layout.addWidget(checkbox)
        layout.setAlignment(Qt.AlignCenter)
        return widget
class BridgeOverallAssessment(BridgeComponentBase):
    """全桥技术状况评定"""
    def __init__(self):
        super().__init__()
        self.component_type = "全桥技术状况评定"
        self.init_ui()
    def init_ui(self):
        """初始化UI"""
        # 创建主布局
        widget = QWidget()
        self.setCentralWidget(widget)
        main_layout = QVBoxLayout(widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(5, 5, 5, 5)
        # 添加标题
        title = QLabel("全桥技术状况评定")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                padding: 4px 8px;
                margin: 1px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                        stop:0 #f6f7fa, stop:1 #e6e6e6);
                border-radius: 2px;
                color: #333333;
            }
        """)
        main_layout.addWidget(title)
        # 创建上部区域容器（包含第一个表格）
        top_container = QWidget()
        top_layout = QVBoxLayout(top_container)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        # 创建评定表格
        self.assessment_table = QTableWidget()
        self.assessment_table.setColumnCount(7)
        self.assessment_table.setHorizontalHeaderLabels([
            '桥梁部位', '结构状况\n指数BSI', '结构状况\n等级', '技术状况\n指数BCI',
            '权重', '技术状况指\n数BCI', '全桥技术状况等级'
        ])
        # 设置表格样式
        self.assessment_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 2px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 2px;
            }
        """)
        # 设置列宽
        column_widths = [150, 100, 100, 150, 100, 150, 158]
        for col, width in enumerate(column_widths):
            self.assessment_table.setColumnWidth(col, width)
        # 设置表格属性
        self.assessment_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.assessment_table.verticalHeader().setVisible(False)
        self.assessment_table.verticalHeader().setDefaultSectionSize(25)
        # 添加固定行
        self.assessment_table.setRowCount(3)
        row_labels = ['桥面系', '上部结构', '下部结构']
        # 初始化表格内容
        for row, label in enumerate(row_labels):
            # 添加桥梁部位
            item = QTableWidgetItem(label)
            item.setTextAlignment(Qt.AlignCenter)
            self.assessment_table.setItem(row, 0, item)
            # 初始化其他单元格
            for col in range(1, 7):
                item = QTableWidgetItem('0.00' if col in [1, 3, 4] else '')
                item.setTextAlignment(Qt.AlignCenter)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.assessment_table.setItem(row, col, item)
        # 设置默认权重
        self.assessment_table.item(0, 4).setText('0.10')  # 桥面系权重
        self.assessment_table.item(1, 4).setText('0.45')  # 上部结构权重
        self.assessment_table.item(2, 4).setText('0.45')  # 下部结构权重
        # 合并最后两列的单元格
        self.assessment_table.setSpan(0, 5, 3, 1)  # 合并第6列（技术状况指数BCI）
        self.assessment_table.setSpan(0, 6, 3, 1)  # 合并第7列（全桥技术状况等级）
        # 设置表格的固定高度，只显示实际内容
        header_height = self.assessment_table.horizontalHeader().height()
        row_height = self.assessment_table.verticalHeader().defaultSectionSize()
        total_height = header_height + (row_height * 3) + 2  # 3行内容 + 表头 + 边框
        self.assessment_table.setFixedHeight(total_height)
        top_layout.addWidget(self.assessment_table)
        # 创建下部区域容器（包含标题和第二个表格）
        bottom_container = QWidget()
        bottom_layout = QVBoxLayout(bottom_container)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(0)
        # 添加评定标准标题
        grade_title = QLabel("桥梁技术状况评定标准")
        grade_title.setAlignment(Qt.AlignCenter)
        grade_title.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                padding: 4px 8px;
                margin: 0px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                        stop:0 #f6f7fa, stop:1 #e6e6e6);
                border-radius: 2px;
                color: #333333;
            }
        """)
        bottom_layout.addWidget(grade_title)
        # 添加评定等级说明表
        grade_table = QTableWidget()
        grade_table.setColumnCount(4)
        grade_table.setRowCount(5)
        grade_table.setHorizontalHeaderLabels(['等级', '状态', 'BCI 范围', '养护对策'])
        # 设置表格样式
        grade_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
                margin-top: 0px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 2px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 4px;
            }
        """)
        # 设置列宽
        grade_table.setColumnWidth(0, 80)
        grade_table.setColumnWidth(1, 100)
        grade_table.setColumnWidth(2, 150)
        grade_table.setColumnWidth(3, 578)
        # 设置表格数据
        grade_data = [
            ('A级', '完好', '[90, 100]', '日常保养'),
            ('B级', '良好', '[80, 90)', '保养小修'),
            ('C级', '合格', '[66, 80)', '针对性小修或中修工程'),
            ('D级', '不合格', '[50, 66)', '检测评估后进行中修、大修或加固工程'),
            ('E级', '危险', '[0, 50)', '检测评估后进行大修、加固或改扩建工程')
        ]
        for row, (grade, state, bci_range, strategy) in enumerate(grade_data):
            # 设置等级
            item = QTableWidgetItem(grade)
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            grade_table.setItem(row, 0, item)
            # 设置状态
            item = QTableWidgetItem(state)
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            grade_table.setItem(row, 1, item)
            # 设置BCI范围
            item = QTableWidgetItem(bci_range)
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            grade_table.setItem(row, 2, item)
            # 设置养护对策
            item = QTableWidgetItem(strategy)
            item.setTextAlignment(Qt.AlignCenter)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            grade_table.setItem(row, 3, item)
        # 设置表格高度
        row_height = 35
        header_height = 25
        total_height = row_height * grade_table.rowCount() + header_height
        grade_table.setFixedHeight(total_height)
        # 隐藏垂直表头
        grade_table.verticalHeader().setVisible(False)
        # 设置水平表头对齐方式
        grade_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        bottom_layout.addWidget(grade_table)
        # 创建一个QVBoxLayout来放置两个容器，并设置它们之间的间距
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(5)  # 减少两个容器之间的间距
        content_layout.addWidget(top_container)
        content_layout.addWidget(bottom_container)
        # 创建一个QWidget来容纳content_layout
        content_widget = QWidget()
        content_widget.setLayout(content_layout)
        # 将content_widget添加到主布局中
        main_layout.addWidget(content_widget)
        # 添加一个弹性空间，使内容居上
        main_layout.addStretch(1)
    def update_data_from_components(self, deck_system, super_structure, sub_structure):
        """从其他模块更新数据"""
        components = [
            (0, deck_system, '桥面系'),
            (1, super_structure, '上部结构'),
            (2, sub_structure, '下部结构')
        ]
        for row, component, name in components:
            # 初始化默认值
            bsi = float('100.00')
            bci = float('100.00')
            if name == '桥面系':
                # 桥面系直接从主表格获取数据
                if hasattr(component, 'table') and component.table:
                    # 获取BSI(m)值
                    bsi_item = component.table.item(0, 8)  # BSI(m)列
                    if bsi_item and bsi_item.text():
                        try:
                            bsi = float(bsi_item.text())
                        except ValueError:
                            bsi = float('100.00')
                    # 获取BCI(m)值
                    bci_item = component.table.item(0, 9)  # BCI(m)列
                    if bci_item and bci_item.text():
                        try:
                            bci = float(bci_item.text())
                        except ValueError:
                            bci = float('100.00')
            else:
                # 上部结构和下部结构从汇总表格获取数据
                if hasattr(component, 'summary_table'):
                    # 获取BSI值和BCI值（在汇总表格的第一行）
                    if component.summary_table.rowCount() > 0:
                        # 获取BCI值（第3列）
                        bci_item = component.summary_table.item(0, 2)
                        if bci_item and bci_item.text():
                            try:
                                bci = float(bci_item.text())
                            except ValueError:
                                bci = float('100.00')
                        # 获取BSI值（第4列）
                        bsi_item = component.summary_table.item(0, 3)
                        if bsi_item and bsi_item.text():
                            try:
                                bsi = float(bsi_item.text())
                            except ValueError:
                                bsi = float('100.00')
                    # 特殊处理下部结构
                    if name == '下部结构':
                        # 检查是否有构件
                        if not hasattr(component, 'all_components') or not component.all_components:
                            bsi = float('100.00')
                            bci = float('100.00')
            # 更新表格数据
            self.assessment_table.item(row, 1).setText(f"{bsi:.2f}")  # BSI
            self.assessment_table.item(row, 3).setText(f"{bci:.2f}")  # BCI
            # 根据BSI值确定结构状况等级
            grade = self.get_grade_info(bsi, include_description=False)
            self.assessment_table.item(row, 2).setText(grade)
            # 计算技术状况指数BCI
            weight = float(self.assessment_table.item(row, 4).text())
            weighted_bci = bci * weight
            self.assessment_table.item(row, 5).setText(f"{weighted_bci:.2f}")
        # 计算全桥技术状况等级
        self.calculate_overall_grade()
    def calculate_overall_grade(self):
        """计算全桥技术状况等级"""
        try:
            # 计算总的BCI值（所有构件的BCI与权重乘积之和）
            total_bci = 0
            for row in range(3):
                bci = float(self.assessment_table.item(row, 3).text())  # 获取BCI值
                weight = float(self.assessment_table.item(row, 4).text())  # 获取权重
                weighted_bci = bci * weight
                # 更新每行的技术状况指数BCI（第6列）
                self.assessment_table.item(row, 5).setText(f"{weighted_bci:.2f}")
                total_bci += weighted_bci
            # 更新合并后的单元格
            self.assessment_table.item(0, 5).setText(f"{total_bci:.2f}")  # 技术状况指数BCI
            self.assessment_table.item(0, 6).setText(
                self.get_grade_info(total_bci, include_description=True)
            )  # 全桥技术状况等级
        except (ValueError, AttributeError) as e:
            QMessageBox.warning(self, "计算错误", f"计算全桥技术状况等级时出错: {str(e)}")
    def get_grade_info(self, value, include_description=False):
        """根据BSI或BCI值获取等级信息
        Args:
            value: BSI或BCI值
            include_description: 是否包含状态描述
        Returns:
            str: 等级或等级加状态描述
        """
        if 90 <= value <= 100:
            grade, state = 'A', '完好状态'
        elif 80 <= value < 90:
            grade, state = 'B', '良好状态'
        elif 66 <= value < 80:
            grade, state = 'C', '合格状态'
        elif 50 <= value < 66:
            grade, state = 'D', '不合格状态'
        elif 0 <= value < 50:
            grade, state = 'E', '危险状态'
        return f"{grade} ({state})" if include_description else grade
class BridgeEvaluationSystem(QMainWindow):
    """桥梁评估系统主窗口"""
    def __init__(self):
        super().__init__()
        self.init_ui()
    def init_ui(self):
        """初始化主窗口UI"""
        # 使用相对路径设置图标
        icon_path = os.path.join(os.path.dirname(__file__), "UrbanBridge.ico")
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle('城市桥梁评估系统-实腹式拱桥')
        # 创建主布局容器
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        # 创建顶部工具栏区域
        toolbar_widget = QWidget()
        toolbar_layout = QHBoxLayout(toolbar_widget)
        toolbar_layout.setContentsMargins(0, 0, 0, 10)
        # 添加工程名称输入框
        project_label = QLabel("工程名称:")
        self.project_name_input = QLineEdit()
        self.project_name_input.setPlaceholderText("请输入桥梁工程名称")
        self.project_name_input.setMinimumWidth(200)
        # 添加保存和历史数据按钮
        self.save_btn = QPushButton("保存数据")
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
        """)
        # 连接保存按钮的信号
        self.save_btn.clicked.connect(self.save_project_data)
        self.history_btn = QPushButton("历史数据")
        self.history_btn.setStyleSheet("""
            QPushButton {
                background-color: #5cb85c;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #449d44;
            }
        """)
        self.history_btn.clicked.connect(self.show_history_data)
        # 添加导出数据按钮
        self.export_btn = QPushButton("导出数据")
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0ad4e;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #ec971f;
            }
        """)
        self.export_btn.clicked.connect(self.export_project_data)
        # 添加到工具栏布局
        toolbar_layout.addWidget(project_label)
        toolbar_layout.addWidget(self.project_name_input)
        toolbar_layout.addWidget(self.save_btn)
        toolbar_layout.addWidget(self.history_btn)
        toolbar_layout.addWidget(self.export_btn)
        toolbar_layout.addStretch()
        # 添加工具栏到主布局
        main_layout.addWidget(toolbar_widget)
        # 创建选项卡控件
        tab_widget = QTabWidget()
        # 设置选项卡样式
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                background: white;
                border-radius: 3px;
            }
            QTabBar::tab {
                background: #f0f0f0;
                border: 1px solid #ddd;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                padding: 8px 15px;
                min-width: 80px;
                font-size: 14px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background: #4a90e2;
                color: white;
                margin-bottom: -1px;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background: #e1e1e1;
            }
        """)
        # 添加各评估模块
        deck_system = DeckSystem()
        super_structure = SuperStructure()
        sub_structure = SubStructure()
        bridge_overall = BridgeOverallAssessment()
        tab_widget.addTab(deck_system, "桥面系")
        tab_widget.addTab(super_structure, "上部结构")
        tab_widget.addTab(sub_structure, "下部结构")
        tab_widget.addTab(bridge_overall, "全桥技术状况评定")
        # 保存模块引用
        self.deck_system = deck_system
        self.super_structure = super_structure
        self.sub_structure = sub_structure
        self.bridge_overall = bridge_overall
        # 添加选项卡切换事件
        tab_widget.currentChanged.connect(self.on_tab_changed)
        # 添加选项卡到主布局
        main_layout.addWidget(tab_widget)
        # 获取屏幕几何信息并设置窗口位置和大小
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - 942) // 2
        y = (screen.height() - 965) // 2
        self.setGeometry(x, y, 942, 965)
    def on_tab_changed(self, index):
        """选项卡切换时更新全桥评定数据"""
        if index == 3:  # 当切换到全桥技术状况评定标签时
            self.bridge_overall.update_data_from_components(
                self.deck_system,
                self.super_structure,
                self.sub_structure
            )
    def save_project_data(self):
        """保存项目数据"""
        try:
            # 获取项目名称
            project_name = self.project_name_input.text().strip()
            if not project_name:
                msg = QMessageBox()
                msg.setWindowTitle("提示")
                msg.setText("请输入工程名称！")
                msg.setIcon(QMessageBox.Warning)
                msg.setStyleSheet("""
                    QMessageBox {
                        background-color: white;
                    }
                    QMessageBox QLabel {
                        color: #333;
                        font-size: 13px;
                        min-height: 20px;
                        max-width: 180px;
                    }
                    QMessageBox QPushButton {
                        background-color: #1890ff;
                        color: white;
                        border: none;
                        border-radius: 3px;
                        padding: 3px 10px;
                        min-width: 60px;
                        font-size: 12px;
                    }
                    QMessageBox QPushButton:hover {
                        background-color: #40a9ff;
                    }
                    /* 调整按钮容器样式 */
                    QDialogButtonBox {
                        alignment: center;
                    }
                    QMessageBox QDialogButtonBox {
                        button-layout: center;
                    }
                """)
                # 获取QMessageBox的布局
                layout = msg.layout()
                # 调整布局间距
                layout.setContentsMargins(15, 15, 15, 15)
                layout.setSpacing(5)
                # 设置文本标签居中
                label = msg.findChild(QLabel)
                if label:
                    label.setAlignment(Qt.AlignCenter)
                # 设置按钮居中
                button_box = msg.findChild(QDialogButtonBox)
                if button_box:
                    button_box.setCenterButtons(True)
                msg.exec()
                return
            # 创建数据目录（如果不存在）
            data_dir = os.path.join(os.path.dirname(__file__), "project_data", "Solid_belly_arch_bridge")
            os.makedirs(data_dir, exist_ok=True)
            # 检查是否存在同名项目的历史数据
            existing_files = []
            for file in os.listdir(data_dir):
                if file.startswith(project_name + "_") and file.endswith(".json"):
                    existing_files.append(file)
            # 如果存在历史数据，询问用户如何处理
            if existing_files:
                msg = QMessageBox()
                msg.setWindowTitle("项目数据已存在")
                msg.setText(f"检测到项目 '{project_name}' 已有{len(existing_files)}个历史版本：\n" + 
                          "\n".join(existing_files))
                msg.setInformativeText("请选择保存方式：")
                # 添加按钮
                save_new = msg.addButton("保存为新版本", QMessageBox.ActionRole)
                overwrite_latest = msg.addButton("覆盖最新版本", QMessageBox.ActionRole)
                cancel = msg.addButton("取消", QMessageBox.RejectRole)
                # 设置样式
                msg.setStyleSheet("""
                    QMessageBox {
                        background-color: white;
                        border: 1px solid #e0e0e0;
                        border-radius: 4px;
                    }
                    QMessageBox QLabel {
                        color: #333;
                        font-size: 12px;
                        min-width: 380px;
                        padding: 10px;
                    }
                    /* 调整按钮容器样式实现居中 */
                    QMessageBox QDialogButtonBox {
                        alignment: center;
                        padding: 10px;
                    }
                    QPushButton {
                        border: none;
                        padding: 6px 12px;
                        border-radius: 3px;
                        font-size: 12px;
                        margin: 0px 8px;  /* 增加按钮之间的间距 */
                        min-width: 90px;  /* 统一按钮宽度 */
                    }
                    QPushButton[text="保存为新版本"] {
                        background-color: #1890ff;
                        color: white;
                    }
                    QPushButton[text="保存为新版本"]:hover {
                        background-color: #40a9ff;
                    }
                    QPushButton[text="覆盖最新版本"] {
                        background-color: #ff4d4f;
                        color: white;
                    }
                    QPushButton[text="覆盖最新版本"]:hover {
                        background-color: #ff7875;
                    }
                    QPushButton[text="取消"] {
                        background-color: #f0f0f0;
                        color: #333;
                    }
                    QPushButton[text="取消"]:hover {
                        background-color: #d9d9d9;
                    }
                """)
                msg.exec()
                clicked_button = msg.clickedButton()
                if clicked_button == cancel:
                    return
                elif clicked_button == overwrite_latest:
                    # 获取最新的文件（按文件名排序）
                    latest_file = sorted(existing_files)[-1]
                    file_name = latest_file
                elif clicked_button == save_new:  # save_new
                    # 创建新的文件名
                    file_name = f"{project_name}_{time.strftime('%Y%m%d_%H%M%S')}.json"
            else:
                # 如果不存在历史数据，直接创建新文件
                file_name = f"{project_name}_{time.strftime('%Y%m%d_%H%M%S')}.json"
            # 构建项目数据
            project_data = {
                "project_name": project_name,
                "save_time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "deck_system": self.get_deck_system_data(),
                "super_structure": self.get_super_structure_data(),
                "sub_structure": self.get_sub_structure_data(),
                "overall_assessment": self.get_overall_assessment_data()
            }
            # 保存数据
            file_path = os.path.join(data_dir, file_name)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "保存成功", f"项目数据已保存至：\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"保存数据时发生错误：\n{str(e)}")
    def get_deck_system_data(self):
        """获取桥面系数据"""
        deck_data = {
            "table_data": [],
            "summary": {}
        }
        if hasattr(self.deck_system, 'table'):
            # 获取表格数据
            for row in range(self.deck_system.table.rowCount()):
                row_data = {
                    "component": self.get_cell_widget_text(self.deck_system.table, row, 0),
                    "damage_type": self.get_cell_text(self.deck_system.table, row, 1),
                    "damage_degree": self.get_combo_text(self.deck_system.table, row, 2),
                    "deduct_points": self.get_cell_text(self.deck_system.table, row, 3),
                    "weight": self.get_cell_text(self.deck_system.table, row, 4),
                    "product": self.get_cell_text(self.deck_system.table, row, 5),
                    "mdp": self.get_cell_text(self.deck_system.table, row, 6),
                    "component_weight": self.get_cell_text(self.deck_system.table, row, 7),
                    "bsi": self.get_cell_text(self.deck_system.table, row, 8),
                    "bci": self.get_cell_text(self.deck_system.table, row, 9)
                }
                deck_data["table_data"].append(row_data)
        return deck_data
    def get_super_structure_data(self):
        """获取上部结构数据"""
        super_data = {
            "spans": {},
            "summary_data": []
        }
        # 保存当前跨的数据，确保数据是最新的
        if hasattr(self.super_structure, 'save_current_span_data'):
            self.super_structure.save_current_span_data()
        # 处理所有跨的数据
        if hasattr(self.super_structure, 'span_data'):
            for span_number, span_data in self.super_structure.span_data.items():
                # 创建每个跨的数据结构
                super_data["spans"][str(span_number)] = {
                    "table_data": [],
                    "bci": span_data.get("bci", 100.0),
                    "checkbox_states": span_data.get("checkbox_states", {})
                }
                # 处理表格数据
                for row, row_data in enumerate(span_data.get("table_data", [])):
                    # 获取复选框状态
                    checkbox_state = None
                    # 注意：这里的键可能是字符串或整数，需要统一处理
                    row_key = row  # 使用整数作为键
                    if row_key in span_data.get("checkbox_states", {}):
                        is_checked = span_data["checkbox_states"][row_key]
                        # 获取复选框文本
                        checkbox_text = ""
                        if hasattr(self.super_structure, 'table'):
                            cell_widget = self.super_structure.table.cellWidget(row, 0)  # 使用整数行号
                            if cell_widget:
                                checkbox = cell_widget.findChild(QCheckBox)
                                if checkbox:
                                    checkbox_text = checkbox.text()
                        checkbox_state = {
                            "text": checkbox_text,
                            "checked": is_checked
                        }
                    # 创建行数据
                    formatted_row_data = {
                        "damage_type": row_data.get("damage_type", ""),
                        "damage_degree": row_data.get("damage_degree", "无"),
                        "deduct_points": row_data.get("deduct_points", "0"),
                        "weight": row_data.get("weight", "0.00"),
                        "product": row_data.get("product", "0.00"),
                        "sdp": row_data.get("sdp", "0.00"),
                        "component_weight": row_data.get("component_weight", "0.00")
                    }
                    # 添加复选框状态
                    if checkbox_state:
                        formatted_row_data["component"] = checkbox_state
                    super_data["spans"][str(span_number)]["table_data"].append(formatted_row_data)
        # 获取汇总表数据
        if hasattr(self.super_structure, 'summary_table'):
            for row in range(self.super_structure.summary_table.rowCount()):
                row_data = {
                    "span": self.get_cell_text(self.super_structure.summary_table, row, 0),
                    "bci": self.get_cell_text(self.super_structure.summary_table, row, 1),
                    "bci_x": self.get_cell_text(self.super_structure.summary_table, row, 2),
                    "bsi_x": self.get_cell_text(self.super_structure.summary_table, row, 3)
                }
                super_data["summary_data"].append(row_data)
        return super_data

    def get_component_name(self, row, structure_type):
        """根据行号和结构类型获取正确的构件名称"""
        # 桥台构件映射
        abutment_components = {
            range(0, 6): "台帽",
            range(6, 13): "台身",
            range(13, 17): "基础",
            range(17, 20): "拱脚",
            range(20, 24): "耳背翼墙"
        }

        # 桥墩构件映射
        pier_components = {
            range(0, 6): "盖梁",
            range(6, 13): "墩身",
            range(13, 17): "基础",
            range(17, 20): "拱脚"
        }

        # 根据结构类型选择对应的映射
        components_map = abutment_components if structure_type == "桥台" else pier_components

        # 根据行号查找对应的构件名称
        for row_range, component_name in components_map.items():
            if row in row_range:
                return component_name

        # 如果没有找到匹配的构件名称，返回空字符串
        return ""

    def get_sub_structure_data(self):
        """获取下部结构数据"""
        sub_data = {
            "abutment_data": {},
            "pier_data": {},
            "all_components": self.sub_structure.all_components if hasattr(self.sub_structure,
                                                                           'all_components') else [],
            "summary_data": []
        }
        # 保存当前构件的数据，确保数据是最新的
        if hasattr(self.sub_structure, 'save_current_data'):
            self.sub_structure.save_current_data()
        # 处理桥台数据
        if hasattr(self.sub_structure, 'abutment_data'):
            for abutment_name, abutment_data in self.sub_structure.abutment_data.items():
                # 创建桥台数据结构
                sub_data["abutment_data"][abutment_name] = {
                    "table_data": [],
                    "bci": abutment_data.get("bci", 100.0),
                    "checkbox_states": abutment_data.get("checkbox_states", {})
                }
                # 处理表格数据
                for row, row_data in enumerate(abutment_data.get("table_data", [])):
                    # 获取复选框状态
                    checkbox_state = None
                    row_key = row  # 使用整数作为键
                    if row_key in abutment_data.get("checkbox_states", {}):
                        is_checked = abutment_data["checkbox_states"][row_key]
                        # 获取构件名称 - 修改这里，使用辅助函数获取正确的构件名称
                        checkbox_text = self.get_component_name(row, "桥台")
                        checkbox_state = {
                            "text": checkbox_text,
                            "checked": is_checked
                        }
                    # 创建行数据
                    formatted_row_data = {
                        "damage_type": row_data.get("damage_type", ""),
                        "damage_degree": row_data.get("damage_degree", "无"),
                        "deduct_points": row_data.get("deduct_points", "0"),
                        "weight": row_data.get("weight", "0.00"),
                        "product": row_data.get("product", "0.00"),
                        "sdp": row_data.get("sdp", "0.00"),
                        "component_weight": row_data.get("component_weight", "0.00")
                    }
                    # 添加复选框状态
                    if checkbox_state:
                        formatted_row_data["component"] = checkbox_state
                    sub_data["abutment_data"][abutment_name]["table_data"].append(formatted_row_data)
        # 处理桥墩数据
        if hasattr(self.sub_structure, 'pier_data'):
            for pier_name, pier_data in self.sub_structure.pier_data.items():
                # 创建桥墩数据结构
                sub_data["pier_data"][pier_name] = {
                    "table_data": [],
                    "bci": pier_data.get("bci", 100.0),
                    "checkbox_states": pier_data.get("checkbox_states", {})
                }
                # 处理表格数据
                for row, row_data in enumerate(pier_data.get("table_data", [])):
                    # 获取复选框状态
                    checkbox_state = None
                    row_key = row  # 使用整数作为键
                    if row_key in pier_data.get("checkbox_states", {}):
                        is_checked = pier_data["checkbox_states"][row_key]
                        # 获取构件名称 - 修改这里，使用辅助函数获取正确的构件名称
                        checkbox_text = self.get_component_name(row, "桥墩")
                        checkbox_state = {
                            "text": checkbox_text,
                            "checked": is_checked
                        }
                    # 创建行数据
                    formatted_row_data = {
                        "damage_type": row_data.get("damage_type", ""),
                        "damage_degree": row_data.get("damage_degree", "无"),
                        "deduct_points": row_data.get("deduct_points", "0"),
                        "weight": row_data.get("weight", "0.00"),
                        "product": row_data.get("product", "0.00"),
                        "sdp": row_data.get("sdp", "0.00"),
                        "component_weight": row_data.get("component_weight", "0.00")
                    }
                    # 添加复选框状态
                    if checkbox_state:
                        formatted_row_data["component"] = checkbox_state
                    sub_data["pier_data"][pier_name]["table_data"].append(formatted_row_data)
        # 获取汇总表数据
        if hasattr(self.sub_structure, 'summary_table'):
            for row in range(self.sub_structure.summary_table.rowCount()):
                row_data = {
                    "component": self.get_cell_text(self.sub_structure.summary_table, row, 0),
                    "bci_xj": self.get_cell_text(self.sub_structure.summary_table, row, 1),
                    "bci_x": self.get_cell_text(self.sub_structure.summary_table, row, 2),
                    "bsi_x": self.get_cell_text(self.sub_structure.summary_table, row, 3)
                }
                sub_data["summary_data"].append(row_data)
        return sub_data
    def get_overall_assessment_data(self):
        """获取全桥技术状况评定数据"""
        overall_data = {
            "table_data": []
        }
        # 如果表格存在，获取表格数据
        if hasattr(self.bridge_overall, 'assessment_table'):
            table = self.bridge_overall.assessment_table
            
            # 先从各组件更新数据，确保数据是最新的
            if hasattr(self.bridge_overall, 'update_data_from_components'):
                self.bridge_overall.update_data_from_components(
                    self.deck_system,
                    self.super_structure,
                    self.sub_structure
                )
            # 再触发一次计算，确保等级和总分是最新的
            elif hasattr(self.bridge_overall, 'calculate_assessment'):
                self.bridge_overall.calculate_assessment()
                
            # 获取表格中的数据
            for row in range(table.rowCount()):
                row_data = {
                    "component": self.get_cell_text(table, row, 0),
                    "bsi": self.get_cell_text(table, row, 1),
                    "grade": self.get_cell_text(table, row, 2),
                    "bci": self.get_cell_text(table, row, 3),
                    "weight": self.get_cell_text(table, row, 4),
                    "weighted_bci": self.get_cell_text(table, row, 5)
                }
                overall_data["table_data"].append(row_data)
            # 获取全桥状况等级
            if table.rowCount() > 0:
                overall_data["final_grade"] = self.get_cell_text(table, 0, 6)
                # 添加全桥BCI值（技术状况指数BCI）
                overall_data["final_bci"] = self.get_cell_text(table, 0, 5)
        return overall_data
    def get_cell_text(self, table, row, col):
        """获取表格单元格文本"""
        item = table.item(row, col)
        return item.text() if item else ""
    def get_cell_widget_text(self, table, row, col):
        """获取表格单元格部件文本"""
        widget = table.cellWidget(row, col)
        if widget:
            checkbox = widget.findChild(QCheckBox)
            if checkbox:
                return {"text": checkbox.text(), "checked": checkbox.isChecked()}
        return {"text": "", "checked": False} 
    def get_combo_text(self, table, row, col):
        """获取下拉框文本"""
        combo = table.cellWidget(row, col)
        return combo.currentText() if combo else ""
    def show_history_data(self):
        """显示历史数据对话框"""
        dialog = HistoryDataDialog(self)
        dialog.exec()
    def load_project_data(self, data):
        """加载项目数据到UI"""
        try:
            # 设置项目名称
            self.project_name_input.setText(data["project_name"])
            # 加载桥面系数据
            if "deck_system" in data:
                self.load_deck_system_data(data["deck_system"])
            # 加载上部结构数据
            if "super_structure" in data:
                self.load_super_structure_data(data["super_structure"])
            # 加载下部结构数据
            if 'sub_structure' in data and hasattr(self, 'sub_structure'):
                # 直接调用现有的load_sub_structure_data方法
                self.load_sub_structure_data(data['sub_structure'])
                # 确保在加载完成后更新汇总表
                if hasattr(self.sub_structure, 'summary_table'):
                    # 刷新汇总表显示
                    for row in range(self.sub_structure.summary_table.rowCount()):
                        for col in range(self.sub_structure.summary_table.columnCount()):
                            item = self.sub_structure.summary_table.item(row, col)
                            if item:
                                item.setTextAlignment(Qt.AlignCenter)
            # 加载全桥评估数据
            if 'overall_assessment' in data and hasattr(self, 'bridge_overall'):
                self.load_overall_assessment_data(data["overall_assessment"])
            QMessageBox.information(self, "成功", "项目数据加载成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载项目数据失败：\n{str(e)}")
    def load_deck_system_data(self, data):
        """加载桥面系数据"""
        try:
            if not hasattr(self.deck_system, 'table'):
                return
            table = self.deck_system.table
            table_data = data.get("table_data", [])
            # 遍历表格数据
            for row, row_data in enumerate(table_data):
                # 设置复选框状态
                if row < table.rowCount():
                    # 处理复选框
                    checkbox_widget = table.cellWidget(row, 0)
                    if checkbox_widget and isinstance(row_data["component"], dict):
                        checkbox = checkbox_widget.findChild(QCheckBox)
                        if checkbox:
                            checkbox.setChecked(row_data["component"].get("checked", False))
                            checkbox.setText(row_data["component"].get("text", ""))
                    # 设置损坏类型
                    self.set_table_cell(table, row, 1, row_data.get("damage_type", ""))
                    # 设置损坏程度（下拉框）
                    combo = table.cellWidget(row, 2)
                    if combo and isinstance(combo, QComboBox):
                        index = combo.findText(row_data.get("damage_degree", ""))
                        if index >= 0:
                            combo.setCurrentIndex(index)
                    # 设置其他数值列
                    self.set_table_cell(table, row, 3, row_data.get("deduct_points", ""))
                    self.set_table_cell(table, row, 4, row_data.get("weight", ""))
                    self.set_table_cell(table, row, 5, row_data.get("product", ""))
                    self.set_table_cell(table, row, 6, row_data.get("mdp", ""))
                    self.set_table_cell(table, row, 7, row_data.get("component_weight", ""))
                    self.set_table_cell(table, row, 8, row_data.get("bsi", ""))
                    self.set_table_cell(table, row, 9, row_data.get("bci", ""))
        except Exception as e:
            raise Exception(f"加载桥面系数据失败: {str(e)}")
    def load_super_structure_data(self, data):
        """加载上部结构数据"""
        try:
            # 检查是否有跨数据（兼容两种可能的键名）
            span_data_key = "spans" if "spans" in data else "span_data"
            if span_data_key in data:
                # 先断开信号连接，避免触发不必要的事件
                self.super_structure.span_combo.blockSignals(True)
                # 获取所有需要的跨号
                spans_needed = sorted([int(span) for span in data[span_data_key].keys()])
                # 清空现有跨数据
                self.super_structure.span_data = {}
                self.super_structure.span_combo.clear()
                # 添加所有需要的跨
                for span_number in spans_needed:
                    self.super_structure.span_combo.addItem(f"第{span_number}跨")
                    # 获取跨数据
                    span_data = data[span_data_key].get(str(span_number), {})
                    # 格式化数据
                    if span_data:
                        formatted_data = {
                            'table_data': [],
                            'bci': span_data.get('bci', 100.00),
                            'checkbox_states': {}
                        }
                        # 处理表格数据
                        for row, row_data in enumerate(span_data.get("table_data", [])):
                            # 保存行数据
                            formatted_row_data = {
                                'damage_type': row_data.get('damage_type', ''),
                                'damage_degree': row_data.get('damage_degree', '无'),
                                'deduct_points': row_data.get('deduct_points', '0'),
                                'weight': row_data.get('weight', '0.00'),
                                'product': row_data.get('product', '0.00'),
                                'sdp': row_data.get('sdp', '0.00'),
                                'component_weight': row_data.get('component_weight', '0.00')
                            }
                            formatted_data['table_data'].append(formatted_row_data)
                            # 处理复选框状态 - 从component字段获取
                            if 'component' in row_data and isinstance(row_data['component'], dict):
                                formatted_data['checkbox_states'][row] = row_data['component'].get('checked', False)
                        # 保存格式化后的数据
                        self.super_structure.span_data[span_number] = formatted_data
                    else:
                        # 如果没有对应的数据，创建默认数据
                        self.super_structure.span_data[span_number] = {
                            'table_data': [],
                            'bci': 100.00,
                            'checkbox_states': {}
                        }
                # 设置当前跨号为第一个跨
                self.super_structure.current_span = spans_needed[0]
                self.super_structure.span_combo.setCurrentIndex(0)
                # 确保load_span_data方法有skip_clear参数
                if hasattr(self.super_structure, 'load_span_data'):
                    # 检查load_span_data方法的参数
                    import inspect
                    sig = inspect.signature(self.super_structure.load_span_data)
                    if 'skip_clear' in sig.parameters:
                        # 如果有skip_clear参数，使用它
                        self.super_structure.load_span_data(self.super_structure.current_span, skip_clear=True)
                    else:
                        # 如果没有skip_clear参数，需要临时修改方法
                        original_method = self.super_structure.load_span_data
                        # 定义一个新的方法，跳过clear_table_data调用
                        def modified_load_span_data(span_number):
                            if span_number in self.super_structure.span_data:
                                span_data = self.super_structure.span_data[span_number]
                                # 恢复复选框状态
                                for row, is_checked in span_data.get('checkbox_states', {}).items():
                                    cell_widget = self.super_structure.table.cellWidget(row, 0)
                                    if cell_widget:
                                        checkbox = cell_widget.findChild(QCheckBox)
                                        if checkbox:
                                            checkbox.setChecked(is_checked)
                                # 恢复表格数据
                                for row, row_data in enumerate(span_data.get('table_data', [])):
                                    if row < self.super_structure.table.rowCount():
                                        # 设置损坏程度
                                        if self.super_structure.table.cellWidget(row, 2) and 'damage_degree' in row_data:
                                            self.super_structure.table.cellWidget(row, 2).setCurrentText(row_data['damage_degree'])
                                        # 恢复其他数值
                                        columns = {
                                            3: 'deduct_points',
                                            4: 'weight',
                                            5: 'product',
                                            6: 'sdp',
                                            7: 'component_weight'
                                        }
                                        for col, key in columns.items():
                                            if row_data.get(key) and self.super_structure.table.item(row, col):
                                                self.super_structure.table.item(row, col).setText(row_data[key])
                                # 恢复BCI值
                                if 'bci' in span_data and self.super_structure.table.item(0, 8):
                                    self.super_structure.table.item(0, 8).setText(f"{span_data['bci']:.2f}")
                        # 临时替换方法
                        self.super_structure.load_span_data = modified_load_span_data
                        # 调用修改后的方法
                        self.super_structure.load_span_data(self.super_structure.current_span)
                        # 恢复原始方法
                        self.super_structure.load_span_data = original_method
                # 重新连接信号
                self.super_structure.span_combo.blockSignals(False)
                # 更新汇总信息
                if hasattr(self.super_structure, 'update_summary'):
                    self.super_structure.update_summary()
            # 加载汇总表数据
            if hasattr(self.super_structure, 'summary_table'):
                table = self.super_structure.summary_table
                summary_data = data.get("summary_data", [])
                # 清空汇总表
                table.setRowCount(0)
                # 添加汇总数据
                for row, row_data in enumerate(summary_data):
                    table.insertRow(row)
                    self.set_table_cell(table, row, 0, row_data.get("span", ""))
                    self.set_table_cell(table, row, 1, row_data.get("bci", ""))
                    self.set_table_cell(table, row, 2, row_data.get("bci_x", ""))
                    self.set_table_cell(table, row, 3, row_data.get("bsi_x", ""))
                # 合并单元格
                if table.rowCount() > 1:
                    table.setSpan(0, 2, table.rowCount(), 1)
                    table.setSpan(0, 3, table.rowCount(), 1)
        except Exception as e:
            raise Exception(f"加载上部结构数据失败: {str(e)}")
    def load_sub_structure_data(self, data):
        """加载下部结构数据"""
        try:
            # 先清空现有构件
            if hasattr(self.sub_structure, 'all_components'):
                # 清空现有数据
                self.sub_structure.all_components = []
                self.sub_structure.abutment_data = {}
                self.sub_structure.pier_data = {}
                # 恢复桥台数据
                abutment_data = data.get("abutment_data", {})
                for abutment_name, abutment_info in abutment_data.items():
                    # 添加到构件列表
                    self.sub_structure.all_components.append({
                        'number': int(abutment_name),
                        'type': '桥台'
                    })
                    # 创建正确格式的数据结构
                    formatted_data = {
                        'table_data': [],
                        'bci': abutment_info.get('bci', 100.00),
                        'checkbox_states': {}
                    }
                    # 处理表格数据
                    for row, row_data in enumerate(abutment_info.get("table_data", [])):
                        # 保存行数据
                        formatted_row_data = {
                            'damage_type': row_data.get('damage_type', ''),
                            'damage_degree': row_data.get('damage_degree', '无'),
                            'deduct_points': row_data.get('deduct_points', '0'),
                            'weight': row_data.get('weight', '0.00'),
                            'product': row_data.get('product', '0.00'),
                            'sdp': row_data.get('sdp', '0.00'),
                            'component_weight': row_data.get('component_weight', '0.00')
                        }
                        formatted_data['table_data'].append(formatted_row_data)
                        # 处理复选框状态 - 从component字段获取
                        if 'component' in row_data and isinstance(row_data['component'], dict):
                            formatted_data['checkbox_states'][row] = row_data['component'].get('checked', False)
                    # 保存格式化后的数据
                    self.sub_structure.abutment_data[int(abutment_name)] = formatted_data
                # 恢复桥墩数据
                pier_data = data.get("pier_data", {})
                for pier_name, pier_info in pier_data.items():
                    # 添加到构件列表
                    self.sub_structure.all_components.append({
                        'number': int(pier_name),
                        'type': '桥墩'
                    })
                    # 创建正确格式的数据结构
                    formatted_data = {
                        'table_data': [],
                        'bci': pier_info.get('bci', 100.00),
                        'checkbox_states': {}
                    }
                    # 处理表格数据
                    for row, row_data in enumerate(pier_info.get("table_data", [])):
                        # 保存行数据
                        formatted_row_data = {
                            'damage_type': row_data.get('damage_type', ''),
                            'damage_degree': row_data.get('damage_degree', '无'),
                            'deduct_points': row_data.get('deduct_points', '0'),
                            'weight': row_data.get('weight', '0.00'),
                            'product': row_data.get('product', '0.00'),
                            'sdp': row_data.get('sdp', '0.00'),
                            'component_weight': row_data.get('component_weight', '0.00')
                        }
                        formatted_data['table_data'].append(formatted_row_data)
                        # 处理复选框状态 - 从component字段获取
                        if 'component' in row_data and isinstance(row_data['component'], dict):
                            formatted_data['checkbox_states'][row] = row_data['component'].get('checked', False)
                    # 保存格式化后的数据
                    self.sub_structure.pier_data[int(pier_name)] = formatted_data
                # 获取当前选中的结构类型
                current_type = "桥台" if self.sub_structure.abutment_radio.isChecked() else "桥墩"
                # 设置结构类型
                self.sub_structure.structure_type = current_type
                # 更新下拉框（先断开信号连接）
                self.sub_structure.pier_combo.blockSignals(True)
                # 更新下拉框，只显示当前类型的构件
                self.sub_structure.update_pier_combo()
                # 如果有对应类型的构件，选择第一个
                current_data = self.sub_structure.abutment_data if current_type == "桥台" else self.sub_structure.pier_data
                if current_data:
                    valid_keys = [k for k in current_data.keys() if k is not None and isinstance(k, (int, float))]
                    if valid_keys:
                        first_pier = min(valid_keys)
                        self.sub_structure.current_pier = first_pier
                        # 在下拉框中选择对应的项
                        index = self.sub_structure.pier_combo.findText(f"第{first_pier}#{'台' if current_type == '桥台' else '墩'}")
                        if index >= 0:
                            self.sub_structure.pier_combo.setCurrentIndex(index)
                        # 加载数据 - 确保在调用load_structure_data前current_pier已正确设置
                        self.sub_structure.load_structure_data()
                # 重新连接信号
                self.sub_structure.pier_combo.blockSignals(False)
                # 先直接更新汇总表，确保使用正确的BCI值
                self.sub_structure.update_summary()
        except Exception as e:
            raise Exception(f"加载下部结构数据失败: {str(e)}")
    def load_overall_assessment_data(self, data):
        """加载全桥技术状况评定数据"""
        try:
            if not hasattr(self.bridge_overall, 'assessment_table'):
                return
            table = self.bridge_overall.assessment_table
            table_data = data.get("table_data", [])
            # 加载表格数据
            for row, row_data in enumerate(table_data):
                if row < table.rowCount():
                    self.set_table_cell(table, row, 0, row_data.get("component", ""))
                    self.set_table_cell(table, row, 1, row_data.get("bsi", ""))
                    self.set_table_cell(table, row, 2, row_data.get("grade", ""))
                    self.set_table_cell(table, row, 3, row_data.get("bci", ""))
                    self.set_table_cell(table, row, 4, row_data.get("weight", ""))
                    self.set_table_cell(table, row, 5, row_data.get("weighted_bci", ""))
            # 设置全桥状况等级
            if "final_grade" in data and table.rowCount() > 0:
                table.item(0, 6).setText(data["final_grade"])
        except Exception as e:
            raise Exception(f"加载全桥技术状况评定数据失败: {str(e)}")
    def set_table_cell(self, table, row, col, value):
        """设置表格单元格的值"""
        if not table.item(row, col):
            item = QTableWidgetItem()
            table.setItem(row, col, item)
        table.item(row, col).setText(str(value))
        table.item(row, col).setTextAlignment(Qt.AlignCenter)
    def export_project_data(self):
        """导出项目数据为Excel格式"""
        try:
            # 获取项目名称
            project_name = self.project_name_input.text().strip()
            if not project_name:
                QMessageBox.warning(self, "提示", "请输入工程名称！")
                return
                
            # 创建export_data目录（如果不存在）
            export_dir = os.path.join(os.path.dirname(__file__), "export_data", "Solid_belly_arch_bridge")
            os.makedirs(export_dir, exist_ok=True)
            
            # 创建带时间戳的文件名
            timestamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
            filename = f"{project_name}_{timestamp}.xlsx"
            export_path = os.path.join(export_dir, filename)
            
            # 创建工作簿
            wb = Workbook()
            # 删除默认的Sheet工作表
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]
            
            # 导出桥面系数据
            self.export_deck_system_data(wb)
            # 导出上部结构数据
            self.export_super_structure_data(wb)
            # 导出下部结构数据
            self.export_sub_structure_data(wb)
            # 导出全桥技术状况评定数据
            self.export_overall_assessment_data(wb)
            
            # 保存工作簿
            wb.save(export_path)
            
            # 询问用户是否打开文件
            reply = QMessageBox.question(self, "导出成功", 
                                        f"数据已成功导出到：\n{export_path}\n是否立即打开文件？",
                                        QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                # 使用系统默认程序打开文件
                os.startfile(export_path) if os.name == 'nt' else os.system(f"open {export_path}")
                
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出数据时发生错误：\n{str(e)}")
    def export_deck_system_data(self, workbook):
        """导出桥面系数据到Excel工作表"""
        # 创建桥面系工作表
        sheet = workbook.create_sheet("桥面系数据")
        
        # 设置标题
        sheet.merge_cells('A1:K1')
        title_cell = sheet['A1']
        title_cell.value = "表8.1 桥面系技术状况指数BCIₘ评定表"
        title_cell.font = Font(name='Times New Roman', size=10.5)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 设置表头 - 使用Excel公式创建下标效果
        headers = [
            '评价要素', 
            '损坏类型', 
            '损坏程度', 
            '单项扣分\nDPₕᵢ', 
            '权重ωₕᵢ', 
            'DPₕᵢ*ωₕᵢ', 
            '各要素评分', 
            '权重ωₕ', 
            'BSIₘ', 
            'BCIₘ', 
            '备注'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
        
        # 获取桥面系表格数据
        table = self.deck_system.table
        row_count = table.rowCount()
        
        # 当前Excel行索引
        excel_row = 3
        
        # 遍历表格行
        row = 0
        selected_components = []
        
        # 首先收集所有选中的部件
        while row < row_count:
            cell_widget = table.cellWidget(row, 0)
            component_name = ""
            is_checked = False
            
            if cell_widget:
                checkbox = cell_widget.findChild(QCheckBox)
                if checkbox:
                    component_name = checkbox.text()
                    is_checked = checkbox.isChecked()
                    
                    if is_checked:
                        row_span = table.rowSpan(row, 0)
                        damage_types = []
                        
                        # 收集该部件下所有有病害的损坏类型（扣分值>0）
                        for i in range(row_span):
                            current_row = row + i
                            deduct_points = float(table.item(current_row, 3).text() if table.item(current_row, 3) else "0")
                            
                            # 只有扣分值大于0的才视为有病害
                            if deduct_points > 0:
                                damage_type = table.item(current_row, 1).text() if table.item(current_row, 1) else ""
                                combo = table.cellWidget(current_row, 2)
                                damage_degree = combo.currentText() if combo else "无"
                                
                                damage_types.append({
                                    "row": current_row,
                                    "type": damage_type,
                                    "degree": damage_degree,
                                    "deduct_points": deduct_points
                                })
                        
                        # 将部件信息添加到选中部件列表
                        selected_components.append({
                            "name": component_name,
                            "start_row": row,
                            "row_span": row_span,
                            "damage_types": damage_types
                        })
            
            row += table.rowSpan(row, 0)
        
        # 处理选中的部件
        for component in selected_components:
            component_name = component["name"]
            start_row = component["start_row"]
            damage_types = component["damage_types"]
            
            start_excel_row = excel_row
            
            # 如果没有病害（扣分值>0的损坏类型），则添加一行"无"
            if not damage_types:
                sheet.cell(row=excel_row, column=1).value = component_name
                sheet.cell(row=excel_row, column=2).value = "无"
                sheet.cell(row=excel_row, column=3).value = "无"
                sheet.cell(row=excel_row, column=4).value = "0"
                sheet.cell(row=excel_row, column=5).value = "0.00"
                sheet.cell(row=excel_row, column=6).value = "0.00"
                
                # 设置单元格样式
                for col in range(1, 12):
                    cell = sheet.cell(row=excel_row, column=col)
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin')
                    )
                
                excel_row += 1
            else:
                # 处理所有有病害的损坏类型
                for i, damage in enumerate(damage_types):
                    current_row = damage["row"]
                    
                    # 添加部位名称（仅第一行显示）
                    if i == 0:
                        sheet.cell(row=excel_row, column=1).value = component_name
                    
                    # 添加损坏类型和损坏程度
                    sheet.cell(row=excel_row, column=2).value = damage["type"]
                    sheet.cell(row=excel_row, column=3).value = damage["degree"]
                    
                    # 获取扣分值
                    deduct_points = damage["deduct_points"]
                    sheet.cell(row=excel_row, column=4).value = int(deduct_points)
                    
                    # 获取权重和乘积
                    weight_value = float(table.item(current_row, 4).text() if table.item(current_row, 4) else "0")
                    product_value = float(table.item(current_row, 5).text() if table.item(current_row, 5) else "0")
                    
                    # 保持小数点后两位
                    sheet.cell(row=excel_row, column=5).value = round(weight_value, 2)
                    sheet.cell(row=excel_row, column=6).value = round(product_value, 2)
                    
                    # 设置单元格样式
                    for col in range(1, 12):
                        cell = sheet.cell(row=excel_row, column=col)
                        cell.font = Font(name='Times New Roman', size=10.5)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = Border(
                            left=Side(style='thin'), 
                            right=Side(style='thin'), 
                            top=Side(style='thin'), 
                            bottom=Side(style='thin')
                        )
                    
                    # 设置数字格式
                    sheet.cell(row=excel_row, column=5).number_format = '0.00'
                    sheet.cell(row=excel_row, column=6).number_format = '0.00'
                    
                    excel_row += 1
            
            # 合并部位单元格
            if excel_row > start_excel_row + 1:
                sheet.merge_cells(f'A{start_excel_row}:A{excel_row-1}')
            
            # 获取MDP(h)和权重ω(h)
            mdp = float(table.item(start_row, 6).text() if table.item(start_row, 6) else "0.00")
            weight = float(table.item(start_row, 7).text() if table.item(start_row, 7) else "0.00")
            
            # 计算各要素评分 = 100 - MDP(h)
            component_score = 100 - mdp
            
            # 添加各要素评分和权重
            sheet.cell(row=start_excel_row, column=7).value = round(component_score, 2)
            sheet.cell(row=start_excel_row, column=8).value = round(weight, 2)
            
            # 设置数字格式
            sheet.cell(row=start_excel_row, column=7).number_format = '0.00'
            sheet.cell(row=start_excel_row, column=8).number_format = '0.00'
            
            # 合并各要素评分和权重单元格
            if excel_row > start_excel_row + 1:
                sheet.merge_cells(f'G{start_excel_row}:G{excel_row-1}')
                sheet.merge_cells(f'H{start_excel_row}:H{excel_row-1}')
            
            # 设置单元格样式
            for col in range(7, 9):
                cell = sheet.cell(row=start_excel_row, column=col)
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin')
                )
        
        # 获取BSI(m)和BCI(m)
        bsi = float(table.item(0, 8).text() if table.item(0, 8) else "100.00")
        bci = float(table.item(0, 9).text() if table.item(0, 9) else "100.00")
        
        # 添加BSI(m)和BCI(m)
        if excel_row > 3:  # 确保有数据行
            sheet.cell(row=3, column=9).value = round(bsi, 2)
            sheet.cell(row=3, column=10).value = round(bci, 2)
            
            # 设置数字格式
            sheet.cell(row=3, column=9).number_format = '0.00'
            sheet.cell(row=3, column=10).number_format = '0.00'
            
            # 合并BSI(m)和BCI(m)单元格
            if excel_row > 4:
                sheet.merge_cells(f'I3:I{excel_row-1}')
                sheet.merge_cells(f'J3:J{excel_row-1}')
            
            # 设置单元格样式
            for col in range(9, 11):
                cell = sheet.cell(row=3, column=col)
                cell.font = Font(name='Times New Roman', size=10.5)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin')
                )
        
        # 处理备注列 - 合并为一个大单元格并输入"-"
        if excel_row > 3:  # 确保有数据行
            sheet.cell(row=3, column=11).value = "——"
            if excel_row > 4:
                sheet.merge_cells(f'K3:K{excel_row-1}')
            
            cell = sheet.cell(row=3, column=11)
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
        
        # 设置列宽
        column_widths = [12, 20, 10, 12, 8, 10, 10, 8, 8, 8, 8]
        for i, width in enumerate(column_widths, 1):
            sheet.column_dimensions[get_column_letter(i)].width = width
        
        # 设置表格最外侧边框为双线框
        max_row = excel_row - 1
        if max_row >= 3:  # 确保有数据行
            # 设置上边框
            for col in range(1, 12):
                cell = sheet.cell(row=2, column=col)
                cell.border = Border(
                    left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='double'), 
                    bottom=Side(style='thin')
                )
            
            # 设置下边框
            for col in range(1, 12):
                cell = sheet.cell(row=max_row, column=col)
                cell.border = Border(
                    left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='double')
                )
            
            # 设置左边框
            for row in range(2, max_row + 1):
                cell = sheet.cell(row=row, column=1)
                current_border = cell.border
                cell.border = Border(
                    left=Side(style='double'), 
                    right=current_border.right, 
                    top=current_border.top, 
                    bottom=current_border.bottom
                )
            
            # 设置右边框
            for row in range(2, max_row + 1):
                cell = sheet.cell(row=row, column=11)
                current_border = cell.border
                cell.border = Border(
                    left=current_border.left, 
                    right=Side(style='double'), 
                    top=current_border.top, 
                    bottom=current_border.bottom
                )
            
            # 设置四个角落的单元格，确保边框正确连接
            # 左上角
            cell = sheet.cell(row=2, column=1)
            cell.border = Border(
                left=Side(style='double'), 
                right=Side(style='thin'), 
                top=Side(style='double'), 
                bottom=Side(style='thin')
            )
            
            # 右上角
            cell = sheet.cell(row=2, column=11)
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='double'), 
                top=Side(style='double'), 
                bottom=Side(style='thin')
            )
            
            # 左下角
            cell = sheet.cell(row=max_row, column=1)
            cell.border = Border(
                left=Side(style='double'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='double')
            )
            
            # 右下角
            cell = sheet.cell(row=max_row, column=11)
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='double'), 
                top=Side(style='thin'), 
                bottom=Side(style='double')
            )
    def export_super_structure_data(self, workbook):
        """导出上部结构数据到Excel工作表"""
        # 创建上部结构工作表
        sheet = workbook.create_sheet("上部结构数据")
        
        # 设置标题
        sheet.merge_cells('A1:L1')
        title_cell = sheet['A1']
        title_cell.value = "表8.2 上部结构状况指数BCIₛ评定表"
        title_cell.font = Font(name='Times New Roman', size=10.5)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 设置表头 - 使用Excel公式创建下标效果
        headers = [
            '跨号', 
            '构件类型', 
            '损坏类型', 
            '损坏程度', 
            '单项扣分\nDPᵢⱼₖ', 
            '权重ωᵢⱼₖ',
            'DPᵢⱼₖ*ωᵢⱼₖ', 
            '各构件大类\n评分', 
            '权重ωᵢⱼ', 
            'BCIₛᵢ', 
            'BSIₛ', 
            'BCIₛ', 
            '备注'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
        
        # 当前Excel行索引
        excel_row = 3
        
        # 获取上部结构数据
        super_data = self.get_super_structure_data()
        
        # 获取所有跨的数据
        if 'spans' in super_data and super_data['spans']:
            # 按跨号排序
            span_numbers = sorted([int(span) for span in super_data['spans'].keys()])
            
            # 获取BCI(s)和BSI(s)值
            bci_values = []
            for span_num in span_numbers:
                span_data = super_data['spans'][str(span_num)]
                bci = span_data.get('bci', 100.00)
                bci_values.append(bci)
            
            bci_s = sum(bci_values) / len(bci_values) if bci_values else 100.00
            bsi_s = min(bci_values) if bci_values else 100.00
            
            # 记录表格的起始行，用于后续合并单元格
            table_start_row = excel_row
            
            # 对跨进行分组，将相同数据的跨分到一组
            span_groups = []
            processed_spans = set()
            
            for i, span_num in enumerate(span_numbers):
                if span_num in processed_spans:
                    continue
                    
                span_data = super_data['spans'][str(span_num)]
                current_group = [span_num]
                processed_spans.add(span_num)
                
                # 比较其他跨与当前跨的数据是否相同
                for j, other_span in enumerate(span_numbers):
                    if other_span in processed_spans or other_span == span_num:
                        continue
                        
                    other_span_data = super_data['spans'][str(other_span)]
                    
                    # 检查BCI值是否相同
                    if abs(span_data.get('bci', 100.00) - other_span_data.get('bci', 100.00)) > 0.01:
                        continue
                    
                    # 检查表格数据是否相同
                    span_table_data = span_data.get('table_data', [])
                    other_table_data = other_span_data.get('table_data', [])
                    
                    # 简化比较：检查构件类型、损坏类型、损坏程度、扣分值是否相同
                    if self._compare_span_data(span_table_data, other_table_data):
                        current_group.append(other_span)
                        processed_spans.add(other_span)
                
                span_groups.append(current_group)
            
            # 处理每组跨的数据
            for group in span_groups:
                # 使用第一个跨的数据作为代表
                span_num = group[0]
                span_data = super_data['spans'][str(span_num)]
                
                # 保存当前跨的起始行
                span_start_row = excel_row
                
                # 获取当前跨的BCI值
                bci_si = span_data.get('bci', 100.00)
                
                # 获取当前跨的表格数据
                table_data = span_data.get('table_data', [])
                
                # 收集构件和损坏类型
                components = {}
                current_component = None
                
                # 首先找出所有选中的构件
                for row_data in table_data:
                    if 'component' in row_data and row_data['component'].get('checked', True):
                        component_name = row_data['component'].get('text', '')
                        component_weight = float(row_data.get('component_weight', '0.00'))
                        sdp = float(row_data.get('sdp', '0.00'))

                        
                        components[component_name] = {
                            'name': component_name,
                            'damage_types': [],
                            'sdp': sdp,
                            'component_weight': component_weight
                        }
                        current_component = component_name
                    
                    # 如果当前行有损坏类型且扣分值大于0，添加到当前构件
                    damage_type = row_data.get('damage_type', '')
                    damage_degree = row_data.get('damage_degree', '无')
                    deduct_points = float(row_data.get('deduct_points', '0'))
                    weight = float(row_data.get('weight', '0.00'))
                    product = float(row_data.get('product', '0.00'))
                    
                    # 只有当损坏程度不是"无"且扣分值大于0时才添加
                    if current_component and damage_type and damage_degree != '无' and deduct_points > 0:
                        components[current_component]['damage_types'].append({
                            'type': damage_type,
                            'degree': damage_degree,
                            'deduct_points': deduct_points,
                            'weight': weight,
                            'product': product
                        })
                
                # 如果没有选中的构件，添加默认的主梁和横向联系
                if not components:
                    components = {
                        '主拱圈': {'name': '主拱圈', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.50},
                        '拱上构造': {'name': '拱上构造', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.20},
                        '横向联系': {'name': '横向联系', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.30}
                    }
                
                # 处理每个构件
                for component_name, component in components.items():
                    damage_types = component['damage_types']
                    
                    component_start_row = excel_row
                    
                    # 如果没有病害，添加一行"无"
                    if not damage_types:
                        # 添加跨号（合并多个跨号）
                        span_text = self._format_span_group(group)
                        sheet.cell(row=excel_row, column=1).value = span_text
                        
                        # 添加构件类型
                        sheet.cell(row=excel_row, column=2).value = component_name
                        
                        # 添加"无"损坏
                        sheet.cell(row=excel_row, column=3).value = "无"
                        sheet.cell(row=excel_row, column=4).value = "无"
                        sheet.cell(row=excel_row, column=5).value = "0"
                        sheet.cell(row=excel_row, column=6).value = "0.00"
                        sheet.cell(row=excel_row, column=7).value = "0.00"
                        
                        # 设置单元格样式
                        for col in range(1, 14):
                            cell = sheet.cell(row=excel_row, column=col)
                            cell.font = Font(name='Times New Roman', size=10.5)
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = Border(
                                left=Side(style='thin'), 
                                right=Side(style='thin'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thin')
                            )
                        
                        # 添加各构件大类评分（100分）
                        sheet.cell(row=excel_row, column=8).value = "100.00"
                        
                        # 添加权重
                        sheet.cell(row=excel_row, column=9).value = f"{component['component_weight']:.2f}"
                        
                        # 添加备注列的值为"——"
                        sheet.cell(row=excel_row, column=13).value = "——"
                        
                        excel_row += 1
                    else:
                        # 处理所有有病害的损坏类型
                        for i, damage in enumerate(damage_types):
                            # 添加跨号（仅在每构件的第一行显示）
                            if i == 0:
                                span_text = self._format_span_group(group)
                                sheet.cell(row=excel_row, column=1).value = span_text
                            
                            # 添加构件类型（仅在每构件的第一行显示）
                            if i == 0:
                                sheet.cell(row=excel_row, column=2).value = component_name
                            
                            # 添加损坏类型和程度
                            sheet.cell(row=excel_row, column=3).value = damage['type']
                            sheet.cell(row=excel_row, column=4).value = damage['degree']
                            
                            # 添加扣分值和权重
                            sheet.cell(row=excel_row, column=5).value = str(int(damage['deduct_points']))
                            sheet.cell(row=excel_row, column=6).value = f"{damage['weight']:.2f}"
                            sheet.cell(row=excel_row, column=7).value = f"{damage['product']:.2f}"
                            
                            # 设置单元格样式
                            for col in range(1, 14):
                                cell = sheet.cell(row=excel_row, column=col)
                                cell.font = Font(name='Times New Roman', size=10.5)
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                cell.border = Border(
                                    left=Side(style='thin'), 
                                    right=Side(style='thin'), 
                                    top=Side(style='thin'), 
                                    bottom=Side(style='thin')
                                )
                            
                            # 添加备注列的值为"——"
                            sheet.cell(row=excel_row, column=13).value = "——"
                            
                            excel_row += 1
                        
                        # 如果有多行，合并构件类型单元格
                        if len(damage_types) > 1:
                            sheet.merge_cells(start_row=component_start_row, start_column=2, 
                                            end_row=excel_row-1, end_column=2)
                    
                    # 在构件的第一行添加SDP和权重
                    sheet.cell(row=component_start_row, column=8).value = f"{100 - component['sdp']:.2f}"
                    sheet.cell(row=component_start_row, column=9).value = f"{component['component_weight']:.2f}"
                    
                    # 如果有多行，合并SDP和权重单元格
                    if excel_row - 1 > component_start_row:
                        sheet.merge_cells(start_row=component_start_row, start_column=8, 
                                        end_row=excel_row-1, end_column=8)
                        sheet.merge_cells(start_row=component_start_row, start_column=9, 
                                        end_row=excel_row-1, end_column=9)
                
                # 合并跨号单元格
                if excel_row - 1 > span_start_row:
                    sheet.merge_cells(start_row=span_start_row, start_column=1, 
                                    end_row=excel_row-1, end_column=1)
                
                # 添加BCI(si)值
                sheet.cell(row=span_start_row, column=10).value = f"{bci_si:.2f}"
                
                # 如果跨有多行，合并BCI(si)单元格
                if excel_row - 1 > span_start_row:
                    sheet.merge_cells(start_row=span_start_row, start_column=10, 
                                    end_row=excel_row-1, end_column=10)
            
            # 添加BSI(s)和BCI(s)值
            if excel_row > table_start_row:
                # 添加BSI(s)值
                sheet.cell(row=table_start_row, column=11).value = f"{bsi_s:.2f}"
                # 添加BCI(s)值
                sheet.cell(row=table_start_row, column=12).value = f"{bci_s:.2f}"
                
                # 合并BSI(s)和BCI(s)单元格
                sheet.merge_cells(start_row=table_start_row, start_column=11, 
                                end_row=excel_row-1, end_column=11)
                sheet.merge_cells(start_row=table_start_row, start_column=12, 
                                end_row=excel_row-1, end_column=12)
                
                # 合并备注列为一个大单元格
                sheet.merge_cells(start_row=table_start_row, start_column=13, 
                                end_row=excel_row-1, end_column=13)
            
            # 设置表格外边框为双线
            # 获取表格的范围
            table_end_row = excel_row - 1
            table_end_col = 13  # 备注列
            
            # 设置表格外边框为双线
            # 上边框
            for col in range(1, table_end_col + 1):
                cell = sheet.cell(row=2, column=col)  # 表头行
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='double'),
                    bottom=Side(style='thin')
                )
            
            # 下边框
            for col in range(1, table_end_col + 1):
                cell = sheet.cell(row=table_end_row, column=col)
                current_border = cell.border
                cell.border = Border(
                    left=current_border.left,
                    right=current_border.right,
                    top=current_border.top,
                    bottom=Side(style='double')
                )
            
            # 左边框
            for row in range(2, table_end_row + 1):
                cell = sheet.cell(row=row, column=1)
                current_border = cell.border
                cell.border = Border(
                    left=Side(style='double'),
                    right=current_border.right,
                    top=current_border.top,
                    bottom=current_border.bottom
                )
            
            # 右边框
            for row in range(2, table_end_row + 1):
                cell = sheet.cell(row=row, column=table_end_col)
                current_border = cell.border
                cell.border = Border(
                    left=current_border.left,
                    right=Side(style='double'),
                    top=current_border.top,
                    bottom=current_border.bottom
                )
            
            # 调整列宽
            for col in range(1, 14):
                sheet.column_dimensions[get_column_letter(col)].width = 10
            
            # 特别调整一些列的宽度
            sheet.column_dimensions['A'].width = 12  # 跨号（增加宽度以容纳多个跨号）
            sheet.column_dimensions['B'].width = 15  # 构件类型
            sheet.column_dimensions['C'].width = 20  # 损坏类型
            sheet.column_dimensions['D'].width = 10  # 损坏程度
            sheet.column_dimensions['H'].width = 10  # 各构件大类评分
            sheet.column_dimensions['M'].width = 10  # 备注

    def _compare_span_data(self, data1, data2):
        """比较两个跨的数据是否相同"""
        if len(data1) != len(data2):
            return False
        
        # 创建两个跨的构件字典
        components1 = {}
        components2 = {}
        
        # 处理第一个跨的数据
        current_component = None
        for row in data1:
            if 'component' in row and row['component'].get('checked', True):
                component_name = row['component'].get('text', '')
                components1[component_name] = {
                    'damages': [],
                    'weight': float(row.get('component_weight', '0.00')),
                    'sdp': float(row.get('sdp', '0.00'))
                }
                current_component = component_name
            
            damage_type = row.get('damage_type', '')
            damage_degree = row.get('damage_degree', '无')
            deduct_points = float(row.get('deduct_points', '0'))
            
            if current_component and damage_type and damage_degree != '无' and deduct_points > 0:
                components1[current_component]['damages'].append({
                    'type': damage_type,
                    'degree': damage_degree,
                    'points': deduct_points
                })
        
        # 处理第二个跨的数据
        current_component = None
        for row in data2:
            if 'component' in row and row['component'].get('checked', True):
                component_name = row['component'].get('text', '')
                components2[component_name] = {
                    'damages': [],
                    'weight': float(row.get('component_weight', '0.00')),
                    'sdp': float(row.get('sdp', '0.00'))
                }
                current_component = component_name
            
            damage_type = row.get('damage_type', '')
            damage_degree = row.get('damage_degree', '无')
            deduct_points = float(row.get('deduct_points', '0'))
            
            if current_component and damage_type and damage_degree != '无' and deduct_points > 0:
                components2[current_component]['damages'].append({
                    'type': damage_type,
                    'degree': damage_degree,
                    'points': deduct_points
                })
        
        # 比较两个构件字典
        if set(components1.keys()) != set(components2.keys()):
            return False
        
        for component_name, component1 in components1.items():
            component2 = components2[component_name]
            
            # 比较权重和SDP
            if abs(component1['weight'] - component2['weight']) > 0.01 or abs(component1['sdp'] - component2['sdp']) > 0.01:
                return False
            
            # 比较损坏情况
            damages1 = sorted(component1['damages'], key=lambda x: (x['type'], x['degree'], x['points']))
            damages2 = sorted(component2['damages'], key=lambda x: (x['type'], x['degree'], x['points']))
            
            if len(damages1) != len(damages2):
                return False
            
            for i in range(len(damages1)):
                d1 = damages1[i]
                d2 = damages2[i]
                if d1['type'] != d2['type'] or d1['degree'] != d2['degree'] or abs(d1['points'] - d2['points']) > 0.01:
                    return False
        
        return True

    def _format_span_group(self, span_group):
        """格式化跨号组为显示文本"""
        if len(span_group) == 1:
            return f"第{span_group[0]}跨"
        
        # 检查是否是连续的跨号
        span_group.sort()
        if all(span_group[i+1] - span_group[i] == 1 for i in range(len(span_group)-1)):
            # 连续跨号使用范围表示
            return f"第{span_group[0]}-{span_group[-1]}跨"
        else:
            # 非连续跨号使用逗号分隔
            return "、".join([f"第{span}跨" for span in span_group])
    def export_sub_structure_data(self, workbook):
        """导出下部结构数据到Excel工作表"""
        # 创建下部结构工作表
        sheet = workbook.create_sheet("下部结构数据")
        
        # 设置标题
        sheet.merge_cells('A1:L1')
        title_cell = sheet['A1']
        title_cell.value = "表8.3 下部结构状况指数BCIₓ评定表"
        title_cell.font = Font(name='Times New Roman', size=10.5)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 设置表头 - 使用Excel公式创建下标效果
        headers = [
            '墩台号', 
            '构件类型', 
            '损坏类型', 
            '损坏程度', 
            '单项扣分\nDPⱼₖₗ', 
            '权重ωⱼₖₗ',
            'DPⱼₖₗ*ωⱼₖₗ', 
            '各构件大类\n评分', 
            '权重ωⱼₖ', 
            'BCIₓⱼ', 
            'BSIₓ', 
            'BCIₓ', 
            '备注'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
        
        # 当前Excel行索引
        excel_row = 3
        
        # 获取下部结构数据
        sub_data = self.get_sub_structure_data()
        
        # 收集所有BCI值用于后续计算总体指标
        all_bci_values = []
        
        # 记录表格的起始行，用于后续合并单元格
        table_start_row = excel_row
        
        # 独立处理桥台数据
        if 'abutment_data' in sub_data and sub_data['abutment_data']:
            # 按桥台号排序
            abutment_numbers = sorted(sub_data['abutment_data'].keys())
            
            # 对桥台进行分组，将相同数据的桥台分到一组
            abutment_groups = []
            processed_abutments = set()
            
            for abutment_num in abutment_numbers:
                if abutment_num in processed_abutments:
                    continue
                    
                abutment_data = sub_data['abutment_data'][abutment_num]
                current_group = [abutment_num]
                processed_abutments.add(abutment_num)
                
                # 比较其他桥台与当前桥台的数据是否相同
                for other_abutment in abutment_numbers:
                    if other_abutment in processed_abutments or other_abutment == abutment_num:
                        continue
                        
                    other_abutment_data = sub_data['abutment_data'][other_abutment]
                    
                    # 使用辅助函数比较数据
                    if self._compare_abutment_data(abutment_data, other_abutment_data):
                        current_group.append(other_abutment)
                        processed_abutments.add(other_abutment)
                
                abutment_groups.append(current_group)
            
            # 处理每组桥台的数据
            for group in abutment_groups:
                # 使用第一个桥台的数据作为代表
                abutment_num = group[0]
                abutment_data = sub_data['abutment_data'][abutment_num]
                
                # 获取格式化的桥台编号文本
                abutment_text = self._format_abutment_group(group)
                
                # 保存当前桥台的起始行
                abutment_start_row = excel_row
                
                # 获取当前桥台的BCI值
                bci_xi = abutment_data.get('bci', 100.00)
                for _ in group:
                    all_bci_values.append(bci_xi)
                
                # 获取当前桥台的表格数据
                table_data = abutment_data.get('table_data', [])
                
                # 收集构件和损坏类型
                components = {}
                current_component = None
                
                # 首先找出所有选中的构件
                for row_data in table_data:
                    if 'component' in row_data and row_data['component'].get('checked', True):
                        component_name = row_data['component'].get('text', '')
                        component_weight = float(row_data.get('component_weight', '0.00'))
                        sdp = float(row_data.get('sdp', '0.00'))
                        
                        components[component_name] = {
                            'name': component_name,
                            'damage_types': [],
                            'sdp': sdp,
                            'component_weight': component_weight
                        }
                        current_component = component_name
                    
                    # 如果当前行有损坏类型且扣分值大于0，添加到当前构件
                    damage_type = row_data.get('damage_type', '')
                    damage_degree = row_data.get('damage_degree', '无')
                    deduct_points = float(row_data.get('deduct_points', '0'))
                    weight = float(row_data.get('weight', '0.00'))
                    product = float(row_data.get('product', '0.00'))
                    
                    # 只有当损坏程度不是"无"且扣分值大于0时才添加
                    if current_component and damage_type and damage_degree != '无' and deduct_points > 0:
                        components[current_component]['damage_types'].append({
                            'type': damage_type,
                            'degree': damage_degree,
                            'deduct_points': deduct_points,
                            'weight': weight,
                            'product': product
                        })
                
                # 如果没有选中的构件，添加桥台的默认构件
                if not components:
                    components = {
                        '台帽': {'name': '台帽', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.10},
                        '台身': {'name': '台身', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.30},
                        '基础': {'name': '基础', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.35},
                        '拱脚': {'name': '拱脚', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.15},
                        '耳背翼墙': {'name': '耳背翼墙', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.10}
                    }
                
                # 处理每个构件
                for component_name, component in components.items():
                    damage_types = component['damage_types']
                    
                    component_start_row = excel_row
                    
                    # 如果没有病害，添加一行"无"
                    if not damage_types:
                        # 添加墩台号
                        sheet.cell(row=excel_row, column=1).value = abutment_text
                        
                        # 添加构件类型
                        sheet.cell(row=excel_row, column=2).value = component_name
                        
                        # 添加"无"损坏
                        sheet.cell(row=excel_row, column=3).value = "无"
                        sheet.cell(row=excel_row, column=4).value = "无"
                        sheet.cell(row=excel_row, column=5).value = "0"
                        sheet.cell(row=excel_row, column=6).value = "0.00"
                        sheet.cell(row=excel_row, column=7).value = "0.00"
                        
                        # 设置单元格样式
                        for col in range(1, 14):
                            cell = sheet.cell(row=excel_row, column=col)
                            cell.font = Font(name='Times New Roman', size=10.5)
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = Border(
                                left=Side(style='thin'), 
                                right=Side(style='thin'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thin')
                            )
                        
                        # 添加各构件大类评分（100分）
                        sheet.cell(row=excel_row, column=8).value = "100.00"
                        
                        # 添加权重
                        sheet.cell(row=excel_row, column=9).value = f"{component['component_weight']:.2f}"
                        
                        excel_row += 1
                    else:
                        # 处理所有有病害的损坏类型
                        for i, damage in enumerate(damage_types):
                            # 添加墩台号（仅在每构件的第一行显示）
                            if i == 0:
                                sheet.cell(row=excel_row, column=1).value = abutment_text
                            
                            # 添加构件类型（仅在每构件的第一行显示）
                            if i == 0:
                                sheet.cell(row=excel_row, column=2).value = component_name
                            
                            # 添加损坏类型和程度
                            sheet.cell(row=excel_row, column=3).value = damage['type']
                            sheet.cell(row=excel_row, column=4).value = damage['degree']
                            
                            # 添加扣分值和权重
                            sheet.cell(row=excel_row, column=5).value = str(int(damage['deduct_points']))
                            sheet.cell(row=excel_row, column=6).value = f"{damage['weight']:.2f}"
                            sheet.cell(row=excel_row, column=7).value = f"{damage['product']:.2f}"
                            
                            # 设置单元格样式
                            for col in range(1, 14):
                                cell = sheet.cell(row=excel_row, column=col)
                                cell.font = Font(name='Times New Roman', size=10.5)
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                cell.border = Border(
                                    left=Side(style='thin'), 
                                    right=Side(style='thin'), 
                                    top=Side(style='thin'), 
                                    bottom=Side(style='thin')
                                )
                            
                            excel_row += 1
                        
                        # 如果有多行，合并构件类型单元格
                        if len(damage_types) > 1:
                            sheet.merge_cells(start_row=component_start_row, start_column=2, 
                                            end_row=excel_row-1, end_column=2)
                    
                    # 在构件的第一行添加SDP和权重
                    sheet.cell(row=component_start_row, column=8).value = f"{100 - component['sdp']:.2f}"
                    sheet.cell(row=component_start_row, column=9).value = f"{component['component_weight']:.2f}"
                    
                    # 如果有多行，合并SDP和权重单元格
                    if excel_row - 1 > component_start_row:
                        sheet.merge_cells(start_row=component_start_row, start_column=8, 
                                        end_row=excel_row-1, end_column=8)
                        sheet.merge_cells(start_row=component_start_row, start_column=9, 
                                        end_row=excel_row-1, end_column=9)
                
                # 合并墩台号单元格
                if excel_row - 1 > abutment_start_row:
                    sheet.merge_cells(start_row=abutment_start_row, start_column=1, 
                                    end_row=excel_row-1, end_column=1)
                
                # 添加BCI(xi)值
                sheet.cell(row=abutment_start_row, column=10).value = f"{bci_xi:.2f}"
                
                # 如果桥台有多行，合并BCI(xi)单元格
                if excel_row - 1 > abutment_start_row:
                    sheet.merge_cells(start_row=abutment_start_row, start_column=10, 
                                    end_row=excel_row-1, end_column=10)
        
        # 独立处理桥墩数据
        if 'pier_data' in sub_data and sub_data['pier_data']:
            # 按桥墩号排序
            pier_numbers = sorted(sub_data['pier_data'].keys())
            
            # 对桥墩进行分组，将相同数据的桥墩分到一组
            pier_groups = []
            processed_piers = set()
            
            for pier_num in pier_numbers:
                if pier_num in processed_piers:
                    continue
                    
                pier_data = sub_data['pier_data'][pier_num]
                current_group = [pier_num]
                processed_piers.add(pier_num)
                
                # 比较其他桥墩与当前桥墩的数据是否相同
                for other_pier in pier_numbers:
                    if other_pier in processed_piers or other_pier == pier_num:
                        continue
                        
                    other_pier_data = sub_data['pier_data'][other_pier]
                    
                    # 使用辅助函数比较数据
                    if self._compare_pier_data(pier_data, other_pier_data):
                        current_group.append(other_pier)
                        processed_piers.add(other_pier)
                
                pier_groups.append(current_group)
            
            # 处理每组桥墩的数据
            for group in pier_groups:
                # 使用第一个桥墩的数据作为代表
                pier_num = group[0]
                pier_data = sub_data['pier_data'][pier_num]
                
                # 获取格式化的桥墩编号文本
                pier_text = self._format_pier_group(group)
                
                # 保存当前桥墩的起始行
                pier_start_row = excel_row
                
                # 获取当前桥墩的BCI值
                bci_xi = pier_data.get('bci', 100.00)
                for _ in group:
                    all_bci_values.append(bci_xi)  
                
                # 获取当前桥墩的表格数据
                table_data = pier_data.get('table_data', [])
                
                # 收集构件和损坏类型
                components = {}
                current_component = None
                
                # 首先找出所有选中的构件
                for row_data in table_data:
                    if 'component' in row_data and row_data['component'].get('checked', True):
                        component_name = row_data['component'].get('text', '')
                        component_weight = float(row_data.get('component_weight', '0.00'))
                        sdp = float(row_data.get('sdp', '0.00'))
                        
                        # 修正桥墩构件名称
                        if component_name == '台帽':
                            component_name = '盖梁'
                        elif component_name == '台身':
                            component_name = '墩身'
                        # 支座和基础名称保持不变
                        
                        components[component_name] = {
                            'name': component_name,
                            'damage_types': [],
                            'sdp': sdp,
                            'component_weight': component_weight
                        }
                        current_component = component_name
                    
                    # 如果当前行有损坏类型且扣分值大于0，添加到当前构件
                    damage_type = row_data.get('damage_type', '')
                    damage_degree = row_data.get('damage_degree', '无')
                    deduct_points = float(row_data.get('deduct_points', '0'))
                    weight = float(row_data.get('weight', '0.00'))
                    product = float(row_data.get('product', '0.00'))
                    
                    # 只有当损坏程度不是"无"且扣分值大于0时才添加
                    if current_component and damage_type and damage_degree != '无' and deduct_points > 0:
                        components[current_component]['damage_types'].append({
                            'type': damage_type,
                            'degree': damage_degree,
                            'deduct_points': deduct_points,
                            'weight': weight,
                            'product': product
                        })
                
                # 如果没有选中的构件，添加桥墩的默认构件
                if not components:
                    components = {
                        '盖梁': {'name': '盖梁', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.10},
                        '墩身': {'name': '墩身', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.30},
                        '基础': {'name': '基础', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.45},
                        '拱脚': {'name': '拱脚', 'damage_types': [], 'sdp': 0.00, 'component_weight': 0.15}
                    }
                
                # 处理每个构件
                for component_name, component in components.items():
                    damage_types = component['damage_types']
                    
                    component_start_row = excel_row
                    
                    # 如果没有病害，添加一行"无"
                    if not damage_types:
                        # 添加墩台号
                        sheet.cell(row=excel_row, column=1).value = pier_text
                        
                        # 添加构件类型
                        sheet.cell(row=excel_row, column=2).value = component_name
                        
                        # 添加"无"损坏
                        sheet.cell(row=excel_row, column=3).value = "无"
                        sheet.cell(row=excel_row, column=4).value = "无"
                        sheet.cell(row=excel_row, column=5).value = "0"
                        sheet.cell(row=excel_row, column=6).value = "0.00"
                        sheet.cell(row=excel_row, column=7).value = "0.00"
                        
                        # 设置单元格样式
                        for col in range(1, 14):
                            cell = sheet.cell(row=excel_row, column=col)
                            cell.font = Font(name='Times New Roman', size=10.5)
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = Border(
                                left=Side(style='thin'), 
                                right=Side(style='thin'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thin')
                            )
                        
                        # 添加各构件大类评分（100分）
                        sheet.cell(row=excel_row, column=8).value = "100.00"
                        
                        # 添加权重
                        sheet.cell(row=excel_row, column=9).value = f"{component['component_weight']:.2f}"
                        
                        excel_row += 1
                    else:
                        # 处理所有有病害的损坏类型
                        for i, damage in enumerate(damage_types):
                            # 添加墩台号（仅在每构件的第一行显示）
                            if i == 0:
                                sheet.cell(row=excel_row, column=1).value = pier_text
                            
                            # 添加构件类型（仅在每构件的第一行显示）
                            if i == 0:
                                sheet.cell(row=excel_row, column=2).value = component_name
                            
                            # 添加损坏类型和程度
                            sheet.cell(row=excel_row, column=3).value = damage['type']
                            sheet.cell(row=excel_row, column=4).value = damage['degree']
                            
                            # 添加扣分值和权重
                            sheet.cell(row=excel_row, column=5).value = str(int(damage['deduct_points']))
                            sheet.cell(row=excel_row, column=6).value = f"{damage['weight']:.2f}"
                            sheet.cell(row=excel_row, column=7).value = f"{damage['product']:.2f}"
                            
                            # 设置单元格样式
                            for col in range(1, 14):
                                cell = sheet.cell(row=excel_row, column=col)
                                cell.font = Font(name='Times New Roman', size=10.5)
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                cell.border = Border(
                                    left=Side(style='thin'), 
                                    right=Side(style='thin'), 
                                    top=Side(style='thin'), 
                                    bottom=Side(style='thin')
                                )
                            
                            excel_row += 1
                        
                        # 如果有多行，合并构件类型单元格
                        if len(damage_types) > 1:
                            sheet.merge_cells(start_row=component_start_row, start_column=2, 
                                            end_row=excel_row-1, end_column=2)
                    
                    # 在构件的第一行添加SDP和权重
                    sheet.cell(row=component_start_row, column=8).value = f"{100 - component['sdp']:.2f}"
                    sheet.cell(row=component_start_row, column=9).value = f"{component['component_weight']:.2f}"
                    
                    # 如果有多行，合并SDP和权重单元格
                    if excel_row - 1 > component_start_row:
                        sheet.merge_cells(start_row=component_start_row, start_column=8, 
                                        end_row=excel_row-1, end_column=8)
                        sheet.merge_cells(start_row=component_start_row, start_column=9, 
                                        end_row=excel_row-1, end_column=9)
                
                # 合并墩台号单元格
                if excel_row - 1 > pier_start_row:
                    sheet.merge_cells(start_row=pier_start_row, start_column=1, 
                                    end_row=excel_row-1, end_column=1)
                
                # 添加BCI(xi)值
                sheet.cell(row=pier_start_row, column=10).value = f"{bci_xi:.2f}"
                
                # 如果桥墩有多行，合并BCI(xi)单元格
                if excel_row - 1 > pier_start_row:
                    sheet.merge_cells(start_row=pier_start_row, start_column=10, 
                                    end_row=excel_row-1, end_column=10)
        
        # 计算BSI(x)和BCI(x)
        if all_bci_values:
            # BSI(x) = min(BCI(xi))
            bsi_x = min(all_bci_values)
            # BCI(x) = 平均值
            bci_x = sum(all_bci_values) / len(all_bci_values)
        else:
            bci_x = 100.00
            bsi_x = 100.00
        
        # 添加总体BCI和BSI值
        if excel_row > table_start_row:
            # 添加BSI(x)值
            sheet.cell(row=table_start_row, column=11).value = f"{bsi_x:.2f}"
            sheet.merge_cells(start_row=table_start_row, start_column=11, 
                            end_row=excel_row-1, end_column=11)
            
            # 添加BCI(x)值
            sheet.cell(row=table_start_row, column=12).value = f"{bci_x:.2f}"
            sheet.merge_cells(start_row=table_start_row, start_column=12, 
                            end_row=excel_row-1, end_column=12)
            
            # 添加备注列并合并为一个大单元格，值为"——"
            sheet.cell(row=table_start_row, column=13).value = "——"
            sheet.merge_cells(start_row=table_start_row, start_column=13, 
                            end_row=excel_row-1, end_column=13)
        
        # 设置表格外边框为双线
        # 获取表格的范围
        table_end_row = excel_row - 1
        table_end_col = 13  # 备注列
        
        # 设置表格外边框为双线
        # 上边框
        for col in range(1, table_end_col + 1):
            cell = sheet.cell(row=2, column=col)  # 表头行
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='double'),
                bottom=Side(style='thin')
            )
        
        # 下边框
        for col in range(1, table_end_col + 1):
            cell = sheet.cell(row=table_end_row, column=col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=current_border.right,
                top=current_border.top,
                bottom=Side(style='double')
            )
        
        # 左边框
        for row in range(2, table_end_row + 1):
            cell = sheet.cell(row=row, column=1)
            current_border = cell.border
            cell.border = Border(
                left=Side(style='double'),
                right=current_border.right,
                top=current_border.top,
                bottom=current_border.bottom
            )
        
        # 右边框
        for row in range(2, table_end_row + 1):
            cell = sheet.cell(row=row, column=table_end_col)
            current_border = cell.border
            cell.border = Border(
                left=current_border.left,
                right=Side(style='double'),
                top=current_border.top,
                bottom=current_border.bottom
            )
        
        # 调整列宽
        for col in range(1, 14):
            column_letter = get_column_letter(col)
            sheet.column_dimensions[column_letter].width = 12
        # 特别调整一些列的宽度
        sheet.column_dimensions['A'].width = 12  # 墩台号（增加宽度以容纳多个跨号）
        sheet.column_dimensions['B'].width = 15  # 构件类型
        sheet.column_dimensions['C'].width = 20  # 损坏类型
        sheet.column_dimensions['D'].width = 10  # 损坏程度
        sheet.column_dimensions['H'].width = 10  # 各构件大类评分
        sheet.column_dimensions['M'].width = 10  # 备注
    def _compare_abutment_data(self, data1, data2):
        """比较两个桥台的数据是否相同"""
        # 提取两个桥台的病害数据
        damage_data1 = self._extract_damage_data(data1)
        damage_data2 = self._extract_damage_data(data2)
        
        # 如果一个桥台有病害而另一个没有，它们不相同
        if (len(damage_data1) > 0 and len(damage_data2) == 0) or (len(damage_data1) == 0 and len(damage_data2) > 0):
            return False
        
        # 如果两个桥台都有病害，比较病害数据
        if len(damage_data1) > 0 and len(damage_data2) > 0:
            # 按构件名称和损坏类型排序
            damage_data1.sort(key=lambda x: (x.get('component', ''), x.get('damage_type', '')))
            damage_data2.sort(key=lambda x: (x.get('component', ''), x.get('damage_type', '')))
            
            # 比较排序后的数据
            if len(damage_data1) != len(damage_data2):
                return False
            
            for i in range(len(damage_data1)):
                item1 = damage_data1[i]
                item2 = damage_data2[i]
                
                # 比较关键字段
                if item1.get('component') != item2.get('component'):
                    return False
                if item1.get('damage_type') != item2.get('damage_type'):
                    return False
                if item1.get('damage_degree') != item2.get('damage_degree'):
                    return False
                if abs(item1.get('deduct_points', 0) - item2.get('deduct_points', 0)) > 0.01:
                    return False
        
        # 比较BCI值
        bci1 = float(data1.get('bci', 100.00))
        bci2 = float(data2.get('bci', 100.00))
        if abs(bci1 - bci2) > 0.01:
            return False
        
        return True

    def _extract_damage_data(self, data):
        """提取桥台/桥墩的病害数据"""
        damage_data = []
        
        for row in data:
            if 'component' in row and isinstance(row['component'], dict) and row['component'].get('checked', True):
                component_name = row['component'].get('text', '')
                damage_type = row.get('damage_type', '')
                damage_degree = row.get('damage_degree', '无')
                deduct_points = float(row.get('deduct_points', '0'))
                
                # 只有当损坏程度不是"无"且扣分值大于0时才添加
                if damage_type and damage_degree != '无' and damage_degree != '完好' and deduct_points > 0:
                    damage_data.append({
                        'component': component_name,
                        'damage_type': damage_type,
                        'damage_degree': damage_degree,
                        'deduct_points': deduct_points
                    })
        
        return damage_data
    
    def _compare_pier_data(self, data1, data2):
        """比较两个桥墩的数据是否相同"""
        # 桥墩的比较逻辑与桥台相同
        return self._compare_abutment_data(data1, data2)
    
    def _format_abutment_group(self, group):
        """格式化桥台编号组"""
        if not group:
            return ""
        
        # 对编号进行排序
        sorted_group = sorted(group)
        
        # 如果只有一个编号，直接返回
        if len(sorted_group) == 1:
            return f"第{sorted_group[0]}#台"
        
        # 查找连续的编号段
        segments = []
        current_segment = [sorted_group[0]]
        
        for i in range(1, len(sorted_group)):
            if sorted_group[i] == sorted_group[i-1] + 1:
                current_segment.append(sorted_group[i])
            else:
                segments.append(current_segment)
                current_segment = [sorted_group[i]]
        
        segments.append(current_segment)
        
        # 格式化输出
        result = []
        for segment in segments:
            if len(segment) == 1:
                result.append(f"第{segment[0]}#台")
            else:
                result.append(f"第{segment[0]}#-{segment[-1]}#台")
        
        return "、".join(result)
    
    def _format_pier_group(self, group):
        """格式化桥墩编号组"""
        if not group:
            return ""
        
        # 对编号进行排序
        sorted_group = sorted(group)
        
        # 如果只有一个编号，直接返回
        if len(sorted_group) == 1:
            return f"第{sorted_group[0]}#墩"
        
        # 查找连续的编号段
        segments = []
        current_segment = [sorted_group[0]]
        
        for i in range(1, len(sorted_group)):
            if sorted_group[i] == sorted_group[i-1] + 1:
                current_segment.append(sorted_group[i])
            else:
                segments.append(current_segment)
                current_segment = [sorted_group[i]]
        
        segments.append(current_segment)
        
        # 格式化输出
        result = []
        for segment in segments:
            if len(segment) == 1:
                result.append(f"第{segment[0]}#墩")
            else:
                result.append(f"第{segment[0]}#-{segment[-1]}#墩")
        
        return "、".join(result)
    def export_overall_assessment_data(self, workbook):
        """导出全桥技术状况数据到Excel工作表"""
        # 创建全桥技术状况工作表
        sheet = workbook.create_sheet("全桥技术状况")
        
        # 设置标题
        sheet.merge_cells('A1:G1')
        title_cell = sheet['A1']
        title_cell.value = "表8.4 全桥技术状况评定表"
        title_cell.font = Font(name='Times New Roman', size=10.5)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 设置表头
        headers = [
            '桥梁部位', 
            '结构状况\n指数BSI', 
            '结构状况\n等级', 
            '技术状况\n指数BCI', 
            '权重', 
            '全桥技术状况\n指数BCI', 
            '全桥技术状况等级'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=2, column=col)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=10.5)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
        
        # 设置列宽
        for col in range(1, 7):
            column_letter = get_column_letter(col)
            sheet.column_dimensions[column_letter].width = 15
        
        # 增加养护对策列的宽度
        sheet.column_dimensions[get_column_letter(7)].width = 25
        
        # 获取全桥技术状况数据
        overall_data = self.get_overall_assessment_data()
        
        # 当前Excel行索引
        excel_row = 3
        
        # 填充表格数据
        if 'table_data' in overall_data and overall_data['table_data']:
            # 计算数据行数
            data_rows = len(overall_data['table_data'])
            
            # 合并技术状况指标BCI列
            sheet.merge_cells(start_row=3, start_column=6, end_row=3 + data_rows - 1, end_column=6)
            
            # 设置技术状况指标BCI值（全桥BCI）
            bci_cell = sheet.cell(row=3, column=6)
            # 确保使用正确的键名获取BCI值
            bci_cell.value = overall_data.get('final_bci', '')
            # 如果final_bci不存在，尝试其他可能的键名
            if not bci_cell.value and 'bci' in overall_data:
                bci_cell.value = overall_data.get('bci', '')
            # 确保值是数字格式
            if isinstance(bci_cell.value, (int, float)):
                bci_cell.value = f"{bci_cell.value:.2f}"
            bci_cell.font = Font(name='Times New Roman', size=10.5)
            bci_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 添加调试输出
            print(f"全桥BCI值: {overall_data.get('final_bci', '未找到')} (类型: {type(overall_data.get('final_bci', ''))})")
            print(f"全桥数据键: {list(overall_data.keys())}")
            
            for row_index, row_data in enumerate(overall_data['table_data']):
                current_row = excel_row + row_index
                
                # 添加桥梁部位
                sheet.cell(row=current_row, column=1).value = row_data.get('component', '')
                
                # 添加结构状况指数BSI
                sheet.cell(row=current_row, column=2).value = row_data.get('bsi', '')
                
                # 添加结构状况等级
                sheet.cell(row=current_row, column=3).value = row_data.get('grade', '')
                
                # 添加技术状况指数BCI
                sheet.cell(row=current_row, column=4).value = row_data.get('bci', '')
                
                # 添加权重
                sheet.cell(row=current_row, column=5).value = row_data.get('weight', '')
                
                # 技术状况指标数BCI已经合并处理
                
                # 设置单元格样式
                for col in range(1, 8):
                    # 跳过已合并的技术状况指标BCI列
                    if col == 6 and row_index > 0:
                        continue
                        
                    cell = sheet.cell(row=current_row, column=col)
                    cell.font = Font(name='Times New Roman', size=10.5)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin')
                    )
            
            # 添加全桥技术状况等级
            if 'final_grade' in overall_data:
                # 合并全桥技术状况等级单元格
                sheet.merge_cells(start_row=3, start_column=7, end_row=excel_row + data_rows - 1, end_column=7)
                
                # 设置全桥技术状况等级
                cell = sheet.cell(row=3, column=7)
                cell.value = overall_data.get('final_grade', '')
                cell.font = Font(name='Times New Roman',size=10.5)  # 移除加粗
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 添加表格外边框（双线）
            for row in range(2, excel_row + data_rows):
                for col in range(1, 8):
                    cell = sheet.cell(row=row, column=col)
                    if row == 1 or row == 2 or col == 1 or col == 7 or row == excel_row + data_rows - 1:
                        # 外边框设置为双线
                        if row == 2:  # 顶部边框
                            cell.border = Border(
                                left=cell.border.left if cell.border else Side(style='thin'),
                                right=cell.border.right if cell.border else Side(style='thin'),
                                top=Side(style='double'),
                                bottom=cell.border.bottom if cell.border else Side(style='thin')
                            )
                        if row == excel_row + data_rows - 1:  # 底部边框
                            cell.border = Border(
                                left=cell.border.left if cell.border else Side(style='thin'),
                                right=cell.border.right if cell.border else Side(style='thin'),
                                top=cell.border.top if cell.border else Side(style='thin'),
                                bottom=Side(style='double')
                            )
                        if col == 1:  # 左侧边框
                            cell.border = Border(
                                left=Side(style='double'),
                                right=cell.border.right if cell.border else Side(style='thin'),
                                top=cell.border.top if cell.border else Side(style='thin'),
                                bottom=cell.border.bottom if cell.border else Side(style='thin')
                            )
                        if col == 7:  # 右侧边框
                            cell.border = Border(
                                left=cell.border.left if cell.border else Side(style='thin'),
                                right=Side(style='double'),
                                top=cell.border.top if cell.border else Side(style='thin'),
                                bottom=cell.border.bottom if cell.border else Side(style='thin')
                            )

class HistoryDataDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        self.load_history_data()
        
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("历史数据")
        self.setMinimumSize(700, 500)
        layout = QVBoxLayout(self)
        
        # 添加搜索框
        search_layout = QHBoxLayout()
        search_label = QLabel("项目名称搜索:")
        search_label.setStyleSheet("""
            QLabel {
                font-size: 13px;
                font-weight: bold;
            }
        """)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入项目名称进行搜索...")
        self.search_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 13px;
            }
            QLineEdit:focus {
                border: 1px solid #4a90e2;
            }
        """)
        self.search_input.textChanged.connect(self.filter_projects)
        
        search_button = QPushButton("搜索")
        search_button.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 6px 15px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
        """)
        search_button.clicked.connect(lambda: self.filter_projects(self.search_input.text()))
        
        clear_button = QPushButton("清除")
        clear_button.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                color: #333;
                border: 1px solid #ddd;
                padding: 6px 15px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        clear_button.clicked.connect(self.clear_search)
        
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input, 1)
        search_layout.addWidget(search_button)
        search_layout.addWidget(clear_button)
        layout.addLayout(search_layout)
        
        # 创建历史数据表格
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['项目名称', '保存时间', '文件名'])
        
        # 设置表格样式
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #ddd;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 6px;
                border: 1px solid #ddd;
                font-weight: bold;
                font-size: 13px;
            }
            QTableWidget::item {
                padding: 8px;
                font-size: 13px;
            }
            QTableWidget::item:selected {
                background-color: #e6f2ff;
                color: #333;
            }
        """)
        
        # 设置表格属性
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        
        # 连接双击信号
        self.table.itemDoubleClicked.connect(self.load_selected_project)
        layout.addWidget(self.table)
        
        # 添加状态标签
        self.status_label = QLabel("共找到 0 个项目")
        self.status_label.setStyleSheet("font-size: 12px; color: #666;")
        layout.addWidget(self.status_label)
        
        # 存储所有项目数据
        self.all_projects = []
        
    def load_history_data(self):
        """加载历史数据"""
        try:
            # 创建主数据目录（如果不存在）
            data_dir = os.path.join(os.path.dirname(__file__), "project_data")
            if not os.path.exists(data_dir):
                os.makedirs(data_dir, exist_ok=True)
                
            # 创建桥梁类型特定的子文件夹
            bridge_type_dir = os.path.join(data_dir, "Solid_belly_arch_bridge")
            if not os.path.exists(bridge_type_dir):
                os.makedirs(bridge_type_dir, exist_ok=True)
                return  # 如果是新创建的文件夹，则没有历史数据
                
            # 获取所有JSON文件
            files = [f for f in os.listdir(bridge_type_dir) if f.endswith('.json')]
            
            # 清空存储的项目数据
            self.all_projects = []
            
            # 按保存时间排序（最新的在前）
            files.sort(reverse=True)
            
            for file in files:
                # 从文件名中提取项目名称（去掉时间戳部分）
                project_name = file.split('_')[0]
                
                # 获取文件修改时间
                file_path = os.path.join(bridge_type_dir, file)
                save_time = time.strftime('%Y-%m-%d %H:%M:%S',
                                        time.localtime(os.path.getmtime(file_path)))
                                        
                # 存储项目数据
                self.all_projects.append({
                    'name': project_name,
                    'time': save_time,
                    'file': file
                })
                
            # 显示所有项目
            self.display_projects(self.all_projects)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载历史数据失败：\n{str(e)}")
            
    def display_projects(self, projects):
        """显示项目列表"""
        self.table.setRowCount(len(projects))
        
        for row, project in enumerate(projects):
            # 添加数据到表格
            for col, key in enumerate(['name', 'time', 'file']):
                item = QTableWidgetItem(project[key])
                item.setTextAlignment(Qt.AlignCenter)  # 设置文本居中对齐
                self.table.setItem(row, col, item)
                
        # 更新状态标签
        self.status_label.setText(f"共找到 {len(projects)} 个项目")
        
    def filter_projects(self, text):
        """根据项目名称筛选项目"""
        if not text:
            # 如果搜索框为空，显示所有项目
            self.display_projects(self.all_projects)
            return
            
        # 筛选匹配的项目
        filtered_projects = [
            project for project in self.all_projects
            if text.lower() in project['name'].lower()
        ]
        
        # 显示筛选后的项目
        self.display_projects(filtered_projects)
        
    def clear_search(self):
        """清除搜索条件"""
        self.search_input.clear()
        self.display_projects(self.all_projects)
        
    def load_selected_project(self, item):
        """加载选中的项目数据"""
        row = item.row()
        file_name = self.table.item(row, 2).text()
        file_path = os.path.join(os.path.dirname(__file__), "project_data", "Solid_belly_arch_bridge", file_name)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            # 调用主窗口的加载数据方法
            self.parent().load_project_data(data)
            self.accept()  # 关闭对话框
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载项目数据失败：\n{str(e)}")
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = BridgeEvaluationSystem()
    window.show()
    sys.exit(app.exec())
