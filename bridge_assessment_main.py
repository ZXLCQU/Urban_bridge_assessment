import sys
import os
import subprocess
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QLabel, QPushButton, QHBoxLayout, QSpacerItem,
                               QSizePolicy, QMessageBox, QGroupBox, QTextEdit, 
                               QScrollArea, QTabWidget, QTextBrowser)
from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QIcon, QFont

# 作者信息常量
AUTHOR_NAME = "Mingtao Zhang"
AUTHOR_EMAIL = "1045811673@qq.com"

# 桥梁类型分类配置
BRIDGE_CATEGORIES = {
    "梁桥": [
        {"name": "混凝土梁桥", "script": "Concrete_beam_bridge.py"},
        {"name": "钢结构梁桥", "script": "Steel_structure_beam_bridge.py"},
        {"name": "混凝土刚构桥", "script": "Concrete_rigid_frame_bridge.py"},
    ],
    "拱桥": [
        {"name": "实腹式拱桥", "script": "Solid_belly_arch_bridge.py"},
        {"name": "空腹式拱桥", "script": "Hollow_belly_arch_bridge.py"},
    ],
    "人行天桥": [
        {"name": "钢结构人行天桥", "script": "Steel_structure_pedestrian_overpass.py"},
        {"name": "混凝土人行天桥", "script": "Concrete_pedestrian_overpass.py"},
    ]
}

class HelpWindow(QMainWindow):
    """帮助文档窗口"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("城市桥梁评估系统 - 帮助文档")
        self.setMinimumSize(900, 700)
        self.center_window()
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), "UrbanBridge.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        self.init_ui()
    
    def center_window(self):
        """将窗口居中显示"""
        screen_geometry = QApplication.primaryScreen().geometry()
        window_geometry = self.geometry()
        x = (screen_geometry.width() - window_geometry.width()) // 2
        y = (screen_geometry.height() - window_geometry.height()) // 2
        self.move(x, y)
    
    def init_ui(self):
        """初始化帮助文档界面"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 设置样式
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
            }
        """)
        
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title = QLabel("城市桥梁评估系统帮助文档")
        title.setAlignment(Qt.AlignCenter)
        title.setFont(QFont("KaiTi", 18, QFont.Bold))
        title.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                padding: 15px;
                background-color: white;
                border-radius: 10px;
                margin-bottom: 20px;
            }
        """)
        layout.addWidget(title)
        
        # 创建选项卡
        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                background-color: white;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #ecf0f1;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font-family: KaiTi;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: 2px solid #3498db;
            }
        """)
        
        # 添加各个帮助页面
        tab_widget.addTab(self.create_overview_tab(), "系统概述")
        tab_widget.addTab(self.create_usage_tab(), "使用指南")
        tab_widget.addTab(self.create_evaluation_tab(), "评估方法")
        tab_widget.addTab(self.create_faq_tab(), "常见问题")
        tab_widget.addTab(self.create_contact_tab(), "联系我们")
        
        layout.addWidget(tab_widget)
    
    def create_overview_tab(self):
        """创建系统概述页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        text_browser = QTextBrowser()
        text_browser.setHtml("""
        <style>
        body { font-family: 'KaiTi', 'Times New Roman', serif; }
        h2, h3, h4 { font-family: 'KaiTi', 'Times New Roman', serif; }
        p, li { font-family: 'KaiTi', 'Times New Roman', serif; }
        </style>
        <h2 style="color: #2c3e50;">系统概述</h2>
        <p style="font-size: 14px; line-height: 1.6;">城市桥梁评估系统是一款专业的城市桥梁技术状况评估软件，基于《城市桥梁养护技术标准》(CJJ99-2017)开发。</p>
        <h3 style="color: #34495e;">主要功能</h3>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li><strong>多类型桥梁支持：</strong>支持混凝土梁桥、钢结构梁桥、混凝土刚构桥、实腹式拱桥、空腹式拱桥、钢结构人行天桥、混凝土人行天桥等桥梁类型</li>
            <li><strong>全面评估体系：</strong>包含桥面系、上部结构、下部结构等部件的详细评估</li>
            <li><strong>智能计算：</strong>自动计算扣分值、权重、MDP、BSI、BCI等关键指标</li>
            <li><strong>数据管理：</strong>支持项目保存、加载、导出Excel报告等功能</li>
            <li><strong>多跨支持：</strong>支持多跨桥梁的独立评估和综合分析</li>
        </ul>
        """)
        
        layout.addWidget(text_browser)
        return widget
    
    def create_usage_tab(self):
        """创建使用指南页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        text_browser = QTextBrowser()
        text_browser.setHtml("""
        <style>
        body { font-family: 'KaiTi', 'Times New Roman', serif; }
        h2, h3, h4 { font-family: 'KaiTi', 'Times New Roman', serif; }
        p, li { font-family: 'KaiTi', 'Times New Roman', serif; }
        </style>
        <h2 style="color: #2c3e50;">使用指南</h2>
        <h3 style="color: #34495e;">1. 启动系统</h3>
        <p style="font-size: 14px; line-height: 1.6;">运行主程序后，系统将显示桥梁类型选择界面，包含三大类别：梁桥、拱桥、人行天桥。</p>
        <h3 style="color: #34495e;">2. 选择桥梁类型</h3>
        <p style="font-size: 14px; line-height: 1.6;">根据待评估桥梁的结构特点，点击相应的桥梁类型按钮进入评估界面。</p>
        <h3 style="color: #34495e;">3. 评估操作流程</h3>
        <ol style="font-size: 14px; line-height: 1.8;">
            <li><strong>选择部件：</strong>在顶部选择要评估的部件（桥面系、上部结构、下部结构）</li>
            <li><strong>输入损坏信息：</strong>在表格中选择相应的损坏程度，系统自动计算扣分值</li>
            <li><strong>多跨处理：</strong>对于多跨桥梁，可使用"添加新跨"功能分别评估各跨</li>
            <li><strong>权重调整：</strong>系统根据构件存在情况自动调整权重</li>
            <li><strong>查看结果：</strong>系统实时计算并显示BSI和BCI指标</li>
        </ol>
        
        <h3 style="color: #34495e;">4. 数据管理</h3>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li><strong>保存项目：</strong>点击"保存项目"按钮，输入项目名称保存当前评估数据</li>
            <li><strong>加载项目：</strong>点击"加载项目"按钮，从历史项目中选择并加载数据</li>
            <li><strong>导出报告：</strong>点击"导出Excel"按钮，生成详细的评估报告</li>
        </ul>
        
        <h3 style="color: #34495e;">5. 注意事项</h3>
        <ul style="font-size: 14px; line-height: 1.8;">
            <li>对于不存在的构件，可取消勾选相应复选框</li>
            <li>建议定期保存评估数据，避免数据丢失</li>
            <li>导出的Excel报告包含完整的计算过程和结果</li>
        </ul>
        """)
        
        layout.addWidget(text_browser)
        return widget

    def create_evaluation_tab(self):
        """创建评估方法页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        text_browser = QTextBrowser()
        text_browser.setHtml("""
        <style>
        body { font-family: 'KaiTi', 'Times New Roman', serif; }
        h2, h3, h4 { font-family: 'KaiTi', 'Times New Roman', serif; }
        p, li { font-family: 'KaiTi', 'Times New Roman', serif; }
        </style>
        <h2 style="color: #2c3e50;">评估方法说明</h2>
        
        <h3 style="color: #34495e;">评估指标体系</h3>
        <p style="font-size: 14px; line-height: 1.6;">系统采用《》城市桥梁养护技术标准(CJJ99-2017)的评估方法，主要包括以下指标：</p>
        
        <div style="background-color: #e8f4fd; padding: 15px; border-radius: 5px; margin: 10px 0;">
            <h4 style="color: #2980b9;">DP(hi) - 扣分值</h4>
            <p style="font-size: 14px;">根据损坏类型和损坏程度确定的基础扣分值，反映单项损坏的严重程度。</p>
            
            <h4 style="color: #2980b9;">ω(hi) - 损坏权重</h4>
            <p style="font-size: 14px;">采用公式：ω(hi) = 3.0×(uhi)³ - 5.5×(uhi)² + 3.5×uhi<br>
            其中 uhi = DP(hi) / ΣDP(hi)，体现各损坏在构件中的相对重要性。</p>
            
            <h4 style="color: #2980b9;">MDP(h) - 构件综合扣分值</h4>
            <p style="font-size: 14px;">MDP(h) = min(max(ΣDP(hi)×ω(hi), max(DP(hi))), 100)<br>
            综合考虑构件内各项损坏的累积效应。</p>
            
            <h4 style="color: #2980b9;">BSI(m) - 桥梁结构指数</h4>
            <p style="font-size: 14px;">BSI(m) = 100 - max(MDP(h))<br>
            反映桥梁最薄弱构件的技术状况。</p>
            
            <h4 style="color: #2980b9;">BCI(m) - 桥梁状况指数</h4>
            <p style="font-size: 14px;">BCI(m) = 100 - Σ(MDP(h) × ω(h))<br>
            综合反映桥梁整体技术状况。</p>
        </div>
        
        <h3 style="color: #34495e;">评估等级划分</h3>
        <table style="border-collapse: collapse; width: 100%; margin: 15px 0;">
            <tr style="background-color: #34495e; color: white;">
                <th style="border: 1px solid #ddd; padding: 10px;">技术状况等级</th>
                <th style="border: 1px solid #ddd; padding: 10px;">BCI范围</th>
                <th style="border: 1px solid #ddd; padding: 10px;">状况描述</th>
            </tr>
            <tr>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">A级</td>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">90-100</td>
                <td style="border: 1px solid #ddd; padding: 8px;">桥梁技术状况完好</td>
            </tr>
            <tr style="background-color: #f8f9fa;">
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">B级</td>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">80-90</td>
                <td style="border: 1px solid #ddd; padding: 8px;">桥梁技术状况良好</td>
            </tr>
            <tr>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">C级</td>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">66-80</td>
                <td style="border: 1px solid #ddd; padding: 8px;">桥梁技术状况合格</td>
            </tr>
            <tr style="background-color: #f8f9fa;">
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">D级</td>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">50-66</td>
                <td style="border: 1px solid #ddd; padding: 8px;">桥梁技术状况不合格</td>
            </tr>
            <tr>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">E级</td>
                <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">0-50</td>
                <td style="border: 1px solid #ddd; padding: 8px;">桥梁技术状况危险</td>
            </tr>
        </table>
        """)
        
        layout.addWidget(text_browser)
        return widget
    
    def create_faq_tab(self):
        """创建常见问题页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        text_browser = QTextBrowser()
        text_browser.setHtml("""
        <style>
        body { font-family: 'KaiTi', 'Times New Roman', serif; }
        h2, h3, h4 { font-family: 'KaiTi', 'Times New Roman', serif; }
        p, li { font-family: 'KaiTi', 'Times New Roman', serif; }
        </style>
        <h2 style="color: #2c3e50;">常见问题</h2>
        
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q1: 某个构件不存在怎么办？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 取消勾选相应构件的复选框，系统会自动调整权重分配，确保总权重为1.0。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q2: 多跨桥梁如何评估？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 使用"添加新跨"功能为每跨单独建立评估表，系统会自动计算各跨的独立指标和全桥综合指标。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q3: 如何理解BSI和BCI的区别？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: BSI反映最薄弱构件的状况，BCI反映整体状况。通常BSI ≤ BCI，BSI更保守，BCI更全面。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q4: 导出的Excel报告包含哪些内容？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 包含完整的评估数据、计算过程、各构件详细信息、最终评估结果和等级判定。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q5: 数据保存在哪里？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 项目数据保存在程序目录下的project_data文件夹中，按桥梁类型分类存储。</p>
        </div>
        
        <div style="margin: 15px 0;">
            <h4 style="color: #e74c3c;">Q6: 如何备份评估数据？</h4>
            <p style="font-size: 14px; margin-left: 20px;">A: 可以直接复制project_data文件夹进行备份，或者将重要项目导出为Excel文件保存。</p>
        </div>
        """)
        
        layout.addWidget(text_browser)
        return widget
    
    def create_contact_tab(self):
        """创建联系我们页面"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        text_browser = QTextBrowser()
        text_browser.setHtml(f"""
        <style>
        body {{ font-family: KaiTi, "Times New Roman", serif; }}
        h2, h3, h4 {{ font-family: KaiTi, "Times New Roman", serif; }}
        p, li {{ font-family: KaiTi, "Times New Roman", serif; }}
        </style>
        
        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 10px; margin: 20px 0;">
            <h3 style="color: #34495e;">开发者信息</h3>
            <p style="font-size: 16px; line-height: 1.8;">
                <strong>作者：</strong>{AUTHOR_NAME}<br>
                <strong>邮箱：</strong>{AUTHOR_EMAIL}<br>
                <strong>版本：</strong>v1.0<br>
                <strong>发布日期：</strong>2025年6月
            </p>
        </div>
        
        <div style="background-color: #e8f4fd; padding: 20px; border-radius: 10px; margin: 20px 0;">
            <h3 style="color: #2980b9;">技术支持</h3>
            <p style="font-size: 14px; line-height: 1.6;">
                如果您在使用过程中遇到问题或有改进建议，欢迎通过邮箱联系作者。
            </p>
        </div>
        
        <div style="background-color: #d5f4e6; padding: 20px; border-radius: 10px; margin: 20px 0;">
            <h3 style="color: #27ae60;">版权声明</h3>
            <p style="font-size: 14px; line-height: 1.6;">
                © 2025 城市桥梁评估系统。本软件基于《城市桥梁养护技术标准》(CJJ99-2017)开发，
                仅供学习和研究使用。如需商业使用，请联系开发者获得授权。
            </p>
        </div>
        """)
        
        layout.addWidget(text_browser)
        return widget


class BridgeSelectionWindow(QMainWindow):
    """桥梁类型选择窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("城市桥梁评估系统")
        self.setMinimumSize(800, 600)
        # 设置窗口居中显示
        self.center_window()
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), "UrbanBridge.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        # 创建主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # 设置窗口样式
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                border-radius: 10px;
            }
        """)
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # 创建包含标题和帮助按钮的顶部布局
        self.create_header_layout(main_layout)
        
        # 添加副标题
        subtitle_label = QLabel("请选择要评估的桥梁类型")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_font = QFont("KaiTi")
        subtitle_font.setPointSize(14)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setStyleSheet("color: #34495e; margin-bottom: 40px; letter-spacing: 1px;")
        main_layout.addWidget(subtitle_label)
        
        # 创建分类按钮布局
        self.create_categorized_buttons(main_layout)
        # 添加底部弹性空间
        main_layout.addItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        # 添加作者信息
        self.add_author_info(main_layout)
    def create_header_layout(self, main_layout):
        """创建包含标题的顶部布局"""
        # 添加标题
        title_label = QLabel("城市桥梁评估系统")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont("KaiTi")
        title_font.setPointSize(24)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 15px; letter-spacing: 2px;")
        main_layout.addWidget(title_label)
        
        # 添加帮助按钮（居中显示）
        help_center_layout = QHBoxLayout()
        help_center_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        
        help_main_button = QPushButton("📖 查看帮助文档")
        help_main_button.setMinimumSize(200, 45)
        help_main_button.setFont(QFont("KaiTi", 12, QFont.Bold))
        help_main_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 22px;
                padding: 10px 20px;
                font-weight: bold;
                letter-spacing: 1px;
            }
            QPushButton:hover {
                background-color: #229954;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
        """)
        help_main_button.clicked.connect(self.show_help)
        help_center_layout.addWidget(help_main_button)
        help_center_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        
        main_layout.addLayout(help_center_layout)
    
    def show_help(self):
        """显示帮助文档窗口"""
        try:
            self.help_window = HelpWindow()
            self.help_window.show()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开帮助文档: {str(e)}")
    def center_window(self):
        """将窗口居中显示在屏幕上"""
        screen_geometry = QApplication.primaryScreen().geometry()
        window_geometry = self.geometry()

        x = (screen_geometry.width() - window_geometry.width()) // 2
        y = (screen_geometry.height() - window_geometry.height()) // 2

        self.move(x, y)

    def create_categorized_buttons(self, main_layout):
        """创建分类的桥梁类型按钮"""
        # 创建水平布局来放置三个分类组
        categories_layout = QHBoxLayout()
        categories_layout.setSpacing(30)
        main_layout.addLayout(categories_layout)
        
        # 为每个分类创建组框
        for category_name, bridges in BRIDGE_CATEGORIES.items():
            group_box = self.create_category_group(category_name, bridges)
            categories_layout.addWidget(group_box)

    def create_category_group(self, category_name, bridges):
        """创建桥梁分类组框"""
        group_box = QGroupBox(category_name)
        group_box.setFont(QFont("KaiTi", 14, QFont.Bold))
        group_box.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
                color: #2c3e50;
                background-color: #f5f5f5;
                border-radius: 5px;
            }
        """)
        
        # 创建组内的垂直布局
        group_layout = QVBoxLayout(group_box)
        group_layout.setSpacing(15)
        group_layout.setContentsMargins(15, 20, 15, 15)
        
        # 为每个桥梁类型创建按钮
        for bridge in bridges:
            button = self.create_bridge_button(bridge["name"], bridge["script"])
            group_layout.addWidget(button)
        
        return group_box

    def create_bridge_button(self, bridge_name, script_name):
        """创建桥梁类型按钮"""
        button = QPushButton(bridge_name)
        button.setMinimumSize(180, 80)
        button.setFont(QFont("KaiTi", 11, QFont.Bold))
        button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }
            QPushButton:hover {
                background-color: #2980b9;
                transform: translateY(-2px);
                transition: all 0.3s ease;
            }
            QPushButton:pressed {
                background-color: #1c6ea4;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
        """)
        button.clicked.connect(lambda: self.run_bridge_script(script_name))
        return button

    def run_bridge_script(self, script_name):
        """运行桥梁评估脚本"""
        try:
            script_path = os.path.join(os.path.dirname(__file__), script_name)
            if not os.path.exists(script_path):
                QMessageBox.critical(self, "错误", f"找不到脚本文件: {script_path}")
                return

            # 添加淡出效果
            self.fade_out_animation = QPropertyAnimation(self, b"windowOpacity")
            self.fade_out_animation.setDuration(300)  # 300毫秒的动画时长
            self.fade_out_animation.setStartValue(1.0)
            self.fade_out_animation.setEndValue(0.0)
            self.fade_out_animation.setEasingCurve(QEasingCurve.InOutQuad)  # 添加缓动曲线使动画更自然

            # 当动画结束时再隐藏窗口并启动子进程
            self.fade_out_animation.finished.connect(lambda: self._run_script_after_fade(script_path))

            # 开始动画
            self.fade_out_animation.start()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"运行脚本时出错: {str(e)}")
            self.show()

    def _run_script_after_fade(self, script_path):
        """淡出动画结束后运行脚本"""
        # 隐藏当前窗口
        self.hide()

        # 使用Python解释器运行脚本
        python_exe = sys.executable
        process = subprocess.Popen([python_exe, script_path])

        # 等待脚本执行完毕后再显示主窗口
        def check_process():
            if process.poll() is not None:
                # 脚本结束，准备显示主窗口
                self.prepare_fade_in()
            else:
                # 每100毫秒检查一次进程状态
                QApplication.instance().processEvents()
                self.check_timer.singleShot(100, check_process)

        # 创建单次计时器
        self.check_timer = QTimer()
        self.check_timer.singleShot(100, check_process)

    def prepare_fade_in(self):
        """准备淡入动画"""
        # 设置窗口透明度为0
        self.setWindowOpacity(0.0)
        # 显示窗口但完全透明
        self.show()

        # 创建淡入动画
        self.fade_in_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_in_animation.setDuration(300)  # 300毫秒的动画时长
        self.fade_in_animation.setStartValue(0.0)
        self.fade_in_animation.setEndValue(1.0)
        self.fade_in_animation.setEasingCurve(QEasingCurve.InOutQuad)  # 添加缓动曲线

        # 开始动画
        self.fade_in_animation.start()

    def add_author_info(self, layout):
        """添加作者信息标签"""
        author_widget = QWidget()
        author_layout = QHBoxLayout(author_widget)
        author_layout.setContentsMargins(5, 5, 5, 5)
        # 创建一个带有渐变背景的容器
        info_container = QWidget()
        info_container.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                          stop:0 #3498db, stop:1 #2c3e50);
                border-radius: 15px;
                padding: 5px;
            }
        """)
        info_layout = QHBoxLayout(info_container)
        info_layout.setContentsMargins(10, 5, 10, 5)
        
        # 创建作者信息标签 - 使用混合字体
        author_label = QLabel()
        author_label.setText(f'<span style="font-family: KaiTi;">作者:</span> <span style="font-family: Times New Roman;">{AUTHOR_NAME}</span>')
        
        email_label = QLabel()
        email_label.setText(f'<span style="font-family: KaiTi;">邮箱:</span> <span style="font-family: Times New Roman;">{AUTHOR_EMAIL}</span>')
        
        # 设置基本样式
        for label in [author_label, email_label]:
            label.setStyleSheet("""
                QLabel {
                    color: white;
                    background-color: transparent;
                    font-size: 10pt;
                    font-weight: bold;
                }
            """)
        
        # 创建垂直布局放置作者和邮箱信息
        author_info_layout = QVBoxLayout()
        author_info_layout.setSpacing(2)
        author_info_layout.addWidget(author_label)
        author_info_layout.addWidget(email_label)
        info_layout.addLayout(author_info_layout)
        
        # 添加版权信息
        copyright_label = QLabel("© 2025 城市桥梁评估系统")
        copyright_label.setStyleSheet("""
            QLabel {
                color: #ecf0f1;
                background-color: transparent;
                font-size: 12px;
            }
        """)
        copyright_label.setFont(QFont("KaiTi"))
        info_layout.addWidget(copyright_label)
        
        # 将信息容器添加到主布局
        author_layout.addStretch()
        author_layout.addWidget(info_container)
        layout.addWidget(author_widget)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BridgeSelectionWindow()
    window.show()
    sys.exit(app.exec())