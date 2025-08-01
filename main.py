import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout
from PyQt5.QtGui import QPalette, QColor
from code_splitter_app import CodeSplitterApp
from dataclean import DataCleanerApp
from excel_formatter import ExcelFormatterApp
from discrepancy_tab import DiscrepancyTab
from balance import Level1BalanceTab
from merge import ExcelMergeApp
from compare_excel_tab import CompareExcelTab

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(" اکسل های لودحسابرسی")
        self.setGeometry(100, 100, 900, 600)

        # رنگ پس‌زمینه ملایم
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f2f2f2;
            }
            QTabWidget::pane {
                border: 1px solid #ccc;
                background: #ffffff;
            }
            QTabBar::tab {
                background: #dfe6e9;
                border: 1px solid #b2bec3;
                padding: 8px;
                min-width: 120px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background: #74b9ff;
                color: white;
            }
            QTabBar::tab:hover {
                background: #a29bfe;
                color: white;
            }
        """)

        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        self.init_tabs()

    def init_tabs(self):
        # تب ادغام اکسل‌ها
        self.tab_widget.addTab(CodeSplitterApp(), "تقسیم کد")

        self.tab_widget.addTab(ExcelMergeApp(), "ادغام اکسل‌ها")


        # تب پاکسازی داده‌ها
        self.clean_data_tab = QWidget()
        self.clean_data_layout = QVBoxLayout()
        self.clean_data_app = DataCleanerApp()
        self.clean_data_layout.addWidget(self.clean_data_app)
        self.clean_data_tab.setLayout(self.clean_data_layout)
        self.tab_widget.addTab(self.clean_data_tab, "پاکسازی داده‌ها")

        # تب فرمت کردن اکسل
        self.excel_format_tab = QWidget()
        self.excel_format_layout = QVBoxLayout()
        self.excel_format_app = ExcelFormatterApp()
        self.excel_format_layout.addWidget(self.excel_format_app)
        self.excel_format_tab.setLayout(self.excel_format_layout)
        self.tab_widget.addTab(self.excel_format_tab, "فرمت فایل اکسل")

        # تب تقسیم کد


        # تب بررسی مغایرت‌ها
        self.tab_widget.addTab(DiscrepancyTab(), "بررسی مغایرت‌ها")

        # تب تراز سطح ۱
        self.tab_widget.addTab(Level1BalanceTab(), "تراز  ")

        # تب مقایسه مرحله‌ای
        self.tab_widget.addTab(CompareExcelTab(), "مقایسه لود مرحله‌ای")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
