from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QLineEdit, QComboBox, QRadioButton, QButtonGroup,
    QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog, QTabWidget,
    QGroupBox, QGridLayout, QStatusBar, QTextEdit
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon, QColor
from checker import ExcelChecker

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("莱州农商银行数据校验工具V3.0T      @FanYuP ")
        self.setGeometry(100, 100, 1000, 700)
        # 设置窗口图标
        self.setWindowIcon(QIcon('app_icon.ico'))
        
        # 设置中文字体
        font = QFont("Microsoft YaHei UI", 9)
        self.setFont(font)

        # 主控件和布局
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout()
        self.central_widget.setLayout(self.main_layout)
        
        # 设置样式表
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QWidget {
                background-color: #f0f0f0;
            }
            QLabel {
                color: #333333;
                font-weight: 500;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
            QLineEdit {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 4px;
                background-color: white;
            }
            QComboBox {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 4px;
                background-color: white;
            }
            QRadioButton {
                color: #333333;
                spacing: 5px;
            }
            QTableWidget {
                border: 1px solid #cccccc;
                border-radius: 4px;
                background-color: white;
                alternate-background-color: #f5f5f5;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 6px;
                border: 1px solid #cccccc;
                font-weight: 600;
            }
            QGroupBox {
                border: 1px solid #cccccc;
                border-radius: 4px;
                margin-top: 10px;
                padding: 10px;
                background-color: #f9f9f9;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #555555;
                font-weight: 600;
            }
        """)

        # 添加标签页控件
        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)

        # 创建双表校验标签页
        self.multi_table_tab = QWidget()
        self.multi_table_layout = QVBoxLayout()
        self.multi_table_tab.setLayout(self.multi_table_layout)

        # 创建单表校验标签页
        self.single_table_tab = QWidget()
        self.single_table_layout = QVBoxLayout()
        self.single_table_tab.setLayout(self.single_table_layout)

        # 创建多表校验标签页（四个表）
        self.four_table_tab = QWidget()
        self.four_table_layout = QVBoxLayout()
        self.four_table_tab.setLayout(self.four_table_layout)

        # 创建校验规则标签页
        self.rules_tab = QWidget()
        self.rules_layout = QVBoxLayout()
        self.rules_tab.setLayout(self.rules_layout)

        # 添加标签页到标签页控件
        self.tab_widget.addTab(self.multi_table_tab, "双表校验")
        self.tab_widget.addTab(self.single_table_tab, "单表校验")
        self.tab_widget.addTab(self.four_table_tab, "多表校验")
        self.tab_widget.addTab(self.rules_tab, "校验规则")

        # 设置双表校验标签页内容
        self.setup_multi_table_ui()

        # 设置单表校验标签页内容
        self.setup_single_table_ui()

        # 设置多表校验标签页内容
        self.setup_four_table_ui()

        # 设置校验规则标签页内容
        self.setup_rules_ui()

        # 初始化校验器
        self.checker = ExcelChecker()
        
        # 添加状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def setup_multi_table_ui(self):
        """设置双表校验标签页的UI"""
        # 文件选择区域
        self.setup_file_selection()

        # 校验功能选择区域
        self.setup_check_options()

        # 结果显示区域
        self.setup_result_display()

        # 添加控件到双表校验标签页
        self.multi_table_layout.addWidget(self.file_group)
        self.multi_table_layout.addWidget(self.options_group)
        self.multi_table_layout.addWidget(self.result_table)
        self.multi_table_layout.setStretch(0, 1)
        self.multi_table_layout.setStretch(1, 2)
        self.multi_table_layout.setStretch(2, 7)

    def setup_file_selection(self):
        """设置文件选择区域"""
        self.file_group = QGroupBox("文件选择")
        self.file_layout = QGridLayout()
        
        # 文件1选择
        self.file1_btn = QPushButton("浏览...")
        self.file1_btn.clicked.connect(lambda: self.select_file(1))
        self.file1_path = QLineEdit()
        self.file1_path.setReadOnly(True)
        self.sheet1_combo = QComboBox()

        # 文件2选择
        self.file2_btn = QPushButton("浏览...")
        self.file2_btn.clicked.connect(lambda: self.select_file(2))
        self.file2_path = QLineEdit()
        self.file2_path.setReadOnly(True)
        self.sheet2_combo = QComboBox()

        # 添加到布局
        self.file_layout.addWidget(QLabel("文件1:"), 0, 0)
        self.file_layout.addWidget(self.file1_btn, 0, 1)
        self.file_layout.addWidget(self.file1_path, 0, 2)
        self.file_layout.addWidget(QLabel("Sheet:"), 0, 3)
        self.file_layout.addWidget(self.sheet1_combo, 0, 4)

        self.file_layout.addWidget(QLabel("文件2:"), 1, 0)
        self.file_layout.addWidget(self.file2_btn, 1, 1)
        self.file_layout.addWidget(self.file2_path, 1, 2)
        self.file_layout.addWidget(QLabel("Sheet:"), 1, 3)
        self.file_layout.addWidget(self.sheet2_combo, 1, 4)
        
        # 设置列拉伸
        self.file_layout.setColumnStretch(2, 1)
        self.file_layout.setColumnStretch(4, 1)
        
        self.file_group.setLayout(self.file_layout)

    def setup_check_options(self):
        """设置校验功能选择区域"""
        self.options_group = QGroupBox("校验功能选择")
        self.options_layout = QGridLayout()

        self.option_group = QButtonGroup()

        # 创建单选按钮
        self.option1 = QRadioButton("1.各项贷款核对(G01和G0107)")
        self.option2 = QRadioButton("2.逾期贷款核对(G01和S6301)")
        self.option3 = QRadioButton("3.逾期贷款(60天以上)核对")
        self.option4 = QRadioButton("4.逾期贷款(90天以上)核对")
        self.option5 = QRadioButton("5.各逾期天数贷款核对(S6301和S7101)")
        self.option6 = QRadioButton("6.个人住房贷款期末余额核对(S67现期与基期)")
        self.option7 = QRadioButton("7.逾期贷款核对1(S67和G01)")
        self.option8 = QRadioButton("8.逾期贷款核对2(S67和G01)")
        self.option9 = QRadioButton("9.逾期贷款核对3(S67和G01)")
        self.option10 = QRadioButton("10.房贷余额占比核对1(S67和G01)")
        self.option11 = QRadioButton("11.房贷余额占比核对2(S67和G01)")

        # 添加到按钮组
        self.option_group.addButton(self.option1, 1)
        self.option_group.addButton(self.option2, 2)
        self.option_group.addButton(self.option3, 3)
        self.option_group.addButton(self.option4, 4)
        self.option_group.addButton(self.option5, 5)
        self.option_group.addButton(self.option6, 6)
        self.option_group.addButton(self.option7, 7)
        self.option_group.addButton(self.option8, 8)
        self.option_group.addButton(self.option9, 9)
        self.option_group.addButton(self.option10, 10)
        self.option_group.addButton(self.option11, 11)

        # 添加到网格布局 (按照用户要求的三列布局: 第一列为1-4号，第二列为5-8，第三列为9-11号功能)
        self.options_layout.addWidget(self.option1, 0, 0)
        self.options_layout.addWidget(self.option2, 1, 0)
        self.options_layout.addWidget(self.option3, 2, 0)
        self.options_layout.addWidget(self.option4, 3, 0)
        self.options_layout.addWidget(self.option5, 0, 1)
        self.options_layout.addWidget(self.option6, 1, 1)
        self.options_layout.addWidget(self.option7, 2, 1)
        self.options_layout.addWidget(self.option8, 3, 1)
        self.options_layout.addWidget(self.option9, 0, 2)
        self.options_layout.addWidget(self.option10, 1, 2)
        self.options_layout.addWidget(self.option11, 2, 2)

        # 添加执行按钮
        self.execute_btn = QPushButton("执行校验")
        self.execute_btn.clicked.connect(self.execute_check)
        self.options_layout.addWidget(self.execute_btn, 4, 0, 1, 3)  # 跨3列
        
        # 设置列拉伸
        self.options_layout.setColumnStretch(0, 1)
        self.options_layout.setColumnStretch(1, 1)
        self.options_layout.setColumnStretch(2, 1)
        
        self.options_group.setLayout(self.options_layout)

    def setup_result_display(self):
        """设置结果显示区域"""
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["项目", "文件1值", "文件2值", "结果"])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.result_table.setAlternatingRowColors(True)

    def select_file(self, file_num):
        """选择文件并加载sheet"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )

        if file_path:
            if file_num == 1:
                self.file1_path.setText(file_path)
                self.load_sheets(file_path, self.sheet1_combo)
                self.status_bar.showMessage(f"已选择文件1: {file_path}")
            else:
                self.file2_path.setText(file_path)
                self.load_sheets(file_path, self.sheet2_combo)
                self.status_bar.showMessage(f"已选择文件2: {file_path}")

    def load_sheets(self, file_path, combo_box):
        """加载Excel文件中的sheet名称"""
        combo_box.clear()
        sheets = self.checker.get_sheet_names(file_path)
        combo_box.addItems(sheets)

    def execute_check(self):
        """执行校验"""
        # 获取选中的校验类型
        option_id = self.option_group.checkedId()
        if option_id == -1:
            self.status_bar.showMessage("请先选择校验功能")
            return

        # 获取文件路径和sheet名称
        file1 = self.file1_path.text()
        sheet1 = self.sheet1_combo.currentText()
        file2 = self.file2_path.text()
        sheet2 = self.sheet2_combo.currentText()

        if not file1 or not file2:
            self.status_bar.showMessage("请先选择两个Excel文件")
            return

        # 执行校验并显示结果
        self.status_bar.showMessage("正在执行校验...")
        results = self.checker.execute_check(option_id, file1, sheet1, file2, sheet2)
        self.display_results(results)
        
        # 统计通过数量
        passed = sum(1 for _, _, _, status in results if status)
        self.status_bar.showMessage(f"校验完成: {passed}/{len(results)} 项通过")

    def display_results(self, results):
        """在表格中显示校验结果"""
        self.result_table.setRowCount(len(results))

        for row, (item, val1, val2, status) in enumerate(results):
            self.result_table.setItem(row, 0, QTableWidgetItem(item))
            self.result_table.setItem(row, 1, QTableWidgetItem(str(val1)))
            self.result_table.setItem(row, 2, QTableWidgetItem(str(val2)))

            result_item = QTableWidgetItem("正常" if status else "异常")
            result_item.setForeground(Qt.green if status else Qt.red)
            self.result_table.setItem(row, 3, result_item)

    def setup_single_table_ui(self):
        """设置单表校验标签页的UI"""
        # 文件选择区域
        self.setup_single_file_selection()

        # 校验功能选择区域
        self.setup_single_check_options()

        # 结果显示区域
        self.setup_single_result_display()

        # 添加控件到单表校验标签页
        self.single_table_layout.addWidget(self.single_file_group)
        self.single_table_layout.addWidget(self.single_options_group)
        self.single_table_layout.addWidget(self.single_result_table)
        self.single_table_layout.setStretch(0, 1)
        self.single_table_layout.setStretch(1, 2)
        self.single_table_layout.setStretch(2, 7)

    def setup_single_file_selection(self):
        """设置单表校验的文件选择区域"""
        self.single_file_group = QGroupBox("文件选择")
        self.single_file_layout = QGridLayout()

        # 文件选择
        self.single_file_btn = QPushButton("浏览...")
        self.single_file_btn.clicked.connect(self.select_single_file)
        self.single_file_path = QLineEdit()
        self.single_file_path.setReadOnly(True)
        self.single_sheet_combo = QComboBox()

        # 添加到布局
        self.single_file_layout.addWidget(QLabel("文件:"), 0, 0)
        self.single_file_layout.addWidget(self.single_file_btn, 0, 1)
        self.single_file_layout.addWidget(self.single_file_path, 0, 2)
        self.single_file_layout.addWidget(QLabel("Sheet:"), 0, 3)
        self.single_file_layout.addWidget(self.single_sheet_combo, 0, 4)
        
        # 设置列拉伸
        self.single_file_layout.setColumnStretch(2, 1)
        self.single_file_layout.setColumnStretch(4, 1)
        
        self.single_file_group.setLayout(self.single_file_layout)

    def setup_single_check_options(self):
        """设置单表校验的功能选择区域"""
        self.single_options_group = QGroupBox("校验功能选择")
        self.single_options_layout = QGridLayout()

        self.single_option_group = QButtonGroup()

        # 创建单选按钮
        self.single_option1 = QRadioButton("1.个人购买商业用房贷款核对1(S67)")
        self.single_option2 = QRadioButton("2.个人购买商业用房贷款核对2(S67)")
        self.single_option3 = QRadioButton("3.基于贷款市场报价利率核对(S67)")
        self.single_option4 = QRadioButton("4.证件号码核对(G1404)")

        # 添加到按钮组
        self.single_option_group.addButton(self.single_option1, 1)
        self.single_option_group.addButton(self.single_option2, 2)
        self.single_option_group.addButton(self.single_option3, 3)
        self.single_option_group.addButton(self.single_option4, 4)

        # 添加到布局 (2列布局)
        self.single_options_layout.addWidget(self.single_option1, 0, 0)
        self.single_options_layout.addWidget(self.single_option2, 0, 1)
        self.single_options_layout.addWidget(self.single_option3, 1, 0)
        self.single_options_layout.addWidget(self.single_option4, 1, 1)

        # 添加执行按钮
        self.single_execute_btn = QPushButton("执行校验")
        self.single_execute_btn.clicked.connect(self.execute_single_check)
        self.single_options_layout.addWidget(self.single_execute_btn, 2, 0, 1, 2)  # 跨2列
        
        # 设置列拉伸
        self.single_options_layout.setColumnStretch(0, 1)
        self.single_options_layout.setColumnStretch(1, 1)
        
        self.single_options_group.setLayout(self.single_options_layout)

    def setup_single_result_display(self):
        """设置单表校验的结果显示区域"""
        self.single_result_table = QTableWidget()
        self.single_result_table.setColumnCount(4)
        self.single_result_table.setHorizontalHeaderLabels(["项目", "单元格值", "期望值", "结果"])
        self.single_result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.single_result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.single_result_table.setAlternatingRowColors(True)

    def select_single_file(self):
        """选择单表校验的文件并加载sheet"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )

        if file_path:
            self.single_file_path.setText(file_path)
            self.load_sheets(file_path, self.single_sheet_combo)
            self.status_bar.showMessage(f"已选择文件: {file_path}")

    def execute_single_check(self):
        """执行单表校验"""
        # 获取选中的校验类型
        option_id = self.single_option_group.checkedId()
        if option_id == -1:
            self.status_bar.showMessage("请先选择校验功能")
            return

        # 获取文件路径和sheet名称
        file_path = self.single_file_path.text()
        sheet_name = self.single_sheet_combo.currentText()

        if not file_path:
            self.status_bar.showMessage("请先选择Excel文件")
            return

        # 执行校验并显示结果
        self.status_bar.showMessage("正在执行校验...")
        results = self.checker.execute_single_check(option_id, file_path, sheet_name)
        self.display_single_results(results)
        
        # 统计通过数量
        passed = sum(1 for _, _, _, status in results if status)
        self.status_bar.showMessage(f"校验完成: {passed}/{len(results)} 项通过")

    def display_single_results(self, results):
        """在表格中显示单表校验结果"""
        self.single_result_table.setRowCount(len(results))

        for row, (item, val1, val2, status) in enumerate(results):
            self.single_result_table.setItem(row, 0, QTableWidgetItem(item))
            self.single_result_table.setItem(row, 1, QTableWidgetItem(str(val1)))
            self.single_result_table.setItem(row, 2, QTableWidgetItem(str(val2)))

            result_item = QTableWidgetItem("正常" if status else "异常")
            result_item.setForeground(Qt.green if status else Qt.red)
            self.single_result_table.setItem(row, 3, result_item)

    def setup_four_table_ui(self):
        """设置多表校验标签页的UI（四个表）"""
        # 文件选择区域
        self.setup_four_file_selection()

        # 校验功能选择区域
        self.setup_four_check_options()

        # 结果显示区域
        self.setup_four_result_display()

        # 添加控件到多表校验标签页
        self.four_table_layout.addWidget(self.four_file_group)
        self.four_table_layout.addWidget(self.four_options_group)
        self.four_table_layout.addWidget(self.four_result_table)
        self.four_table_layout.setStretch(0, 1)
        self.four_table_layout.setStretch(1, 2)
        self.four_table_layout.setStretch(2, 7)

    def setup_four_file_selection(self):
        """设置四个表的文件选择区域"""
        self.four_file_group = QGroupBox("文件选择")
        self.four_file_layout = QGridLayout()
        
        # 文件1选择
        self.four_file1_btn = QPushButton("浏览...")
        self.four_file1_btn.clicked.connect(lambda: self.select_four_file(1))
        self.four_file1_path = QLineEdit()
        self.four_file1_path.setReadOnly(True)
        self.four_sheet1_combo = QComboBox()

        # 文件2选择
        self.four_file2_btn = QPushButton("浏览...")
        self.four_file2_btn.clicked.connect(lambda: self.select_four_file(2))
        self.four_file2_path = QLineEdit()
        self.four_file2_path.setReadOnly(True)
        self.four_sheet2_combo = QComboBox()

        # 文件3选择
        self.four_file3_btn = QPushButton("浏览...")
        self.four_file3_btn.clicked.connect(lambda: self.select_four_file(3))
        self.four_file3_path = QLineEdit()
        self.four_file3_path.setReadOnly(True)
        self.four_sheet3_combo = QComboBox()

        # 文件4选择
        self.four_file4_btn = QPushButton("浏览...")
        self.four_file4_btn.clicked.connect(lambda: self.select_four_file(4))
        self.four_file4_path = QLineEdit()
        self.four_file4_path.setReadOnly(True)
        self.four_sheet4_combo = QComboBox()

        # 添加到布局
        self.four_file_layout.addWidget(QLabel("文件1:"), 0, 0)
        self.four_file_layout.addWidget(self.four_file1_btn, 0, 1)
        self.four_file_layout.addWidget(self.four_file1_path, 0, 2)
        self.four_file_layout.addWidget(QLabel("Sheet:"), 0, 3)
        self.four_file_layout.addWidget(self.four_sheet1_combo, 0, 4)

        self.four_file_layout.addWidget(QLabel("文件2:"), 1, 0)
        self.four_file_layout.addWidget(self.four_file2_btn, 1, 1)
        self.four_file_layout.addWidget(self.four_file2_path, 1, 2)
        self.four_file_layout.addWidget(QLabel("Sheet:"), 1, 3)
        self.four_file_layout.addWidget(self.four_sheet2_combo, 1, 4)

        self.four_file_layout.addWidget(QLabel("文件3:"), 2, 0)
        self.four_file_layout.addWidget(self.four_file3_btn, 2, 1)
        self.four_file_layout.addWidget(self.four_file3_path, 2, 2)
        self.four_file_layout.addWidget(QLabel("Sheet:"), 2, 3)
        self.four_file_layout.addWidget(self.four_sheet3_combo, 2, 4)

        self.four_file_layout.addWidget(QLabel("文件4:"), 3, 0)
        self.four_file_layout.addWidget(self.four_file4_btn, 3, 1)
        self.four_file_layout.addWidget(self.four_file4_path, 3, 2)
        self.four_file_layout.addWidget(QLabel("Sheet:"), 3, 3)
        self.four_file_layout.addWidget(self.four_sheet4_combo, 3, 4)
        
        # 设置列拉伸
        self.four_file_layout.setColumnStretch(2, 1)
        self.four_file_layout.setColumnStretch(4, 1)
        
        self.four_file_group.setLayout(self.four_file_layout)

    def setup_four_check_options(self):
        """设置四个表的校验功能选择区域"""
        self.four_options_group = QGroupBox("校验功能选择")
        self.four_options_layout = QGridLayout()

        self.four_option_group = QButtonGroup()

        # 创建单选按钮
        self.four_option1 = QRadioButton("1.利率异常校验(G01和G04)")
        self.four_option2 = QRadioButton("2.贷款拨备覆盖率校验(G01和G04A)")

        # 添加到按钮组
        self.four_option_group.addButton(self.four_option1, 1)
        self.four_option_group.addButton(self.four_option2, 2)

        # 添加到网格布局
        self.four_options_layout.addWidget(self.four_option1, 0, 0)
        self.four_options_layout.addWidget(self.four_option2, 0, 1)

        # 添加执行按钮
        self.four_execute_btn = QPushButton("执行校验")
        self.four_execute_btn.clicked.connect(self.execute_four_check)
        self.four_options_layout.addWidget(self.four_execute_btn, 1, 0, 1, 2)  # 跨2列
        
        # 设置列拉伸
        self.four_options_layout.setColumnStretch(0, 1)
        self.four_options_layout.setColumnStretch(1, 1)
        
        self.four_options_group.setLayout(self.four_options_layout)

    def setup_four_result_display(self):
        """设置四个表的结果显示区域"""
        self.four_result_table = QTableWidget()
        self.four_result_table.setColumnCount(4)
        self.four_result_table.setHorizontalHeaderLabels(["项目", "计算结果", "标准值", "结果"])
        self.four_result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.four_result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.four_result_table.setAlternatingRowColors(True)

    def select_four_file(self, file_num):
        """选择四个表中的文件并加载sheet"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )

        if file_path:
            if file_num == 1:
                self.four_file1_path.setText(file_path)
                self.load_sheets(file_path, self.four_sheet1_combo)
                self.status_bar.showMessage(f"已选择文件1: {file_path}")
            elif file_num == 2:
                self.four_file2_path.setText(file_path)
                self.load_sheets(file_path, self.four_sheet2_combo)
                self.status_bar.showMessage(f"已选择文件2: {file_path}")
            elif file_num == 3:
                self.four_file3_path.setText(file_path)
                self.load_sheets(file_path, self.four_sheet3_combo)
                self.status_bar.showMessage(f"已选择文件3: {file_path}")
            elif file_num == 4:
                self.four_file4_path.setText(file_path)
                self.load_sheets(file_path, self.four_sheet4_combo)
                self.status_bar.showMessage(f"已选择文件4: {file_path}")

    def execute_four_check(self):
        """执行四个表的校验"""
        # 获取选中的校验类型
        option_id = self.four_option_group.checkedId()
        if option_id == -1:
            self.status_bar.showMessage("请先选择校验功能")
            return

        # 获取文件路径和sheet名称
        file1 = self.four_file1_path.text()
        sheet1 = self.four_sheet1_combo.currentText()
        file2 = self.four_file2_path.text()
        sheet2 = self.four_sheet2_combo.currentText()
        file3 = self.four_file3_path.text()
        sheet3 = self.four_sheet3_combo.currentText()
        file4 = self.four_file4_path.text()
        sheet4 = self.four_sheet4_combo.currentText()

        if not file1 or not file2 or not file3 or not file4:
            self.status_bar.showMessage("请先选择四个Excel文件")
            return

        # 执行校验并显示结果
        self.status_bar.showMessage("正在执行校验...")
        results = self.checker.execute_four_check(option_id, file1, sheet1, file2, sheet2, file3, sheet3, file4, sheet4)
        self.display_four_results(results)
        
        # 统计通过数量
        passed = sum(1 for _, _, _, _, status in results if status)
        self.status_bar.showMessage(f"校验完成: {passed}/{len(results)} 项通过")

    def display_four_results(self, results):
        """显示四个表的校验结果"""
        self.four_result_table.setRowCount(0)

        for row_idx, (item, file1_val, file2_val, _, status) in enumerate(results):
            self.four_result_table.insertRow(row_idx)
            
            item_item = QTableWidgetItem(item)
            file1_item = QTableWidgetItem(str(file1_val))
            file2_item = QTableWidgetItem(str(file2_val))
            result_item = QTableWidgetItem("正常" if status else "异常")
            result_item.setForeground(Qt.green if status else Qt.red)
            
            self.four_result_table.setItem(row_idx, 0, item_item)
            self.four_result_table.setItem(row_idx, 1, file1_item)
            self.four_result_table.setItem(row_idx, 2, file2_item)
            self.four_result_table.setItem(row_idx, 3, result_item)

    def setup_rules_ui(self):
        """设置校验规则标签页的UI"""
        # 创建文本编辑框来显示规则
        self.rules_text = QTextEdit()
        self.rules_text.setReadOnly(True)
        self.rules_text.setStyleSheet("QTextEdit {background-color: #f5f5f5; border: none; padding: 10px; font-size: 18px;}")


        # 设置规则内容
        rules_content = """<h2>双表校验规则</h2>
<ol>
<li>各项贷款核对(G01和G0107)：G01中各项贷款应该与G0107中各项贷款核对一致</li>
<li>逾期贷款核对(G01和S6301)：G01中逾期贷款应大于等于S6301中1.3逾期贷款</li>
<li>逾期贷款(60天以上)核对：逾期60天以上贷款大于等于s6301中60天以上贷款</li>
<li>逾期贷款(90天以上)核对：逾期90天以上贷款大于等于s6301中90天以上贷款</li>
<li>各逾期天数贷款核对(S6301和S7101)：S6301各逾期天数贷款和应与S7101各逾期天数贷款数相等</li>
<li>个人住房贷款期末余额核对(S67现期与基期)：S6701表中1.4个人住房贷款现期余额等于基期余额+当月新发放贷款金额-当月收回金额</li>
<li>逾期贷款核对1(S67和G01)：S6701融资风险监测表中贷款逾期情况逾期贷款与G0102逾期贷款核对</li>
<li>逾期贷款核对2(S67和G01)：S6701融资风险监测表中逾期60天以上贷款与G0102逾期贷款核对</li>
<li>逾期贷款核对3(S67和G01)：S6701融资风险监测表中逾期90天数以上贷款与G0102逾期贷款核对</li>
<li>房贷余额占比核对1(S67和G01)：S6701融资风险监测表中房地产贷款合计除以G01中各项贷款小于百分之二十</li>
<li>房贷余额占比核对2(S67和G01)：S6701融资风险监测表中个人住房贷款除以G01中各项贷款小于百分之十五</li>
</ol>

<h2>单表校验规则</h2>
<ol>
<li>个人购买商业用房贷款核对1(S67)：S6701融资风险监测表中 1.3.1.2 个人购买商业用房贷款期末余额=正常类+关注类+不良类</li>
<li>个人购买商业用房贷款核对2(S67)：S6701融资风险监测表中 1.3.1.2 个人购买商业用房贷款期末余额= 1.3.2.1 LTV≤30%+1.3.2.2 30%＜LTV≤40%+1.3.2.3 40%＜LTV≤50%+ 1.3.2.4 LTV﹥50%+ 1.3.2.5 非抵押方式</li>
<li>基于贷款市场报价利率核对(S67)：S6701融资风险监测表中 基于贷款市场报价利率(LPR) 当月新发放金额=1.4.4.2.1R＜LPR+ 1.4.4.2.2R=LPR+1.4.4.2.3LPR＜R＜LPR+60bp+1.4.4.2.4R=LPR+60bp+ 1.4.4.2.5R﹥LPR+60bp的当月新发放金额</li>
<li>证件号码核对(G1404)：G1404表中客户代码为18位</li>
</ol>

<h2>多表校验规则</h2>
<ol>
<li>利率异常校验：G01中各项贷款余额增长大于百分之二十同时G04中贷款利息收入增长大于百分之五，校验结果异常</li>
<li>贷款拨备覆盖率校验：G01中不良贷款余额增长同时1.1实际计提的贷款损失准备余额除以2.1不良贷款余额的得到的贷款拨备覆盖率提升，校验结果异常</li>
</ol>"""

        self.rules_text.setHtml(rules_content)

        # 添加到布局
        self.rules_layout.addWidget(self.rules_text)
