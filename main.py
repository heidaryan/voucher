import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget,
    QVBoxLayout, QPushButton, QFileDialog, QMessageBox,
    QLabel, QComboBox, QTextEdit
)
from code_splitter_app import CodeSplitterApp
from dataclean import DataCleanerApp
from excel_formatter import ExcelFormatterApp
from discrepancy_tab import DiscrepancyTab
from balance import Level1BalanceTab
from merge import ExcelMergeApp
from compare_excel_tab import CompareExcelTab


class CodeLengthNormalizer(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self.levels = {}

        self.layout = QVBoxLayout()

        self.load_button = QPushButton(" انتخاب فایل ریز اسناد  ")
        self.load_button.clicked.connect(self.load_excel)
        self.layout.addWidget(self.load_button)

        self.column_count_label = QLabel("تعداد ستون‌ها (1 تا 4):")
        self.layout.addWidget(self.column_count_label)

        self.column_count_dropdown = QComboBox()
        self.column_count_dropdown.addItems([str(i) for i in range(1, 5)])
        self.layout.addWidget(self.column_count_dropdown)

        self.column_combos = []
        self.additional_column_combo = QComboBox()

        self.confirm_button = QPushButton("تایید انتخاب ستون‌ها")
        self.confirm_button.clicked.connect(self.prepare_columns)
        self.layout.addWidget(self.confirm_button)

        self.process_button = QPushButton("اجرای نرمال‌سازی و ترکیب")
        self.process_button.clicked.connect(self.process_data)
        self.layout.addWidget(self.process_button)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.layout.addWidget(self.output_text)

        self.setLayout(self.layout)

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.output_text.setText("فایل با موفقیت بارگذاری شد.")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در خواندن فایل:\n{e}")

    def prepare_columns(self):
        if self.df is None:
            QMessageBox.warning(self, "خطا", "لطفاً ابتدا یک فایل بارگذاری کنید.")
            return

        # پاک‌سازی انتخاب‌های قبلی
        for combo in self.column_combos:
            self.layout.removeWidget(combo)
            combo.deleteLater()
        self.column_combos = []

        try:
            num_columns = int(self.column_count_dropdown.currentText())
        except ValueError:
            QMessageBox.warning(self, "خطا", "تعداد ستون‌ها نامعتبر است.")
            return

        for _ in range(num_columns):
            combo = QComboBox()
            combo.addItems(["- انتخاب کنید -"] + list(self.df.columns))
            self.column_combos.append(combo)
            self.layout.addWidget(combo)

        self.additional_column_combo = QComboBox()
        self.additional_column_combo.addItems(["- انتخاب کنید -"] + list(self.df.columns))
        self.layout.addWidget(QLabel("ستون مقایسه:"))
        self.layout.addWidget(self.additional_column_combo)

    def normalize_lengths(self):
        for combo in self.column_combos:
            col = combo.currentText()
            if col and col != "- انتخاب کنید -":
                max_len = self.df[col].astype(str).str.len().max()
                self.levels[col] = max_len

        for col, max_len in self.levels.items():
            self.df[col] = self.df[col].apply(
                lambda x: str(int(x)).strip().zfill(max_len) if pd.notna(x) and str(x).strip() != '' else ''
            )

    def concatenate_columns(self):
        def concatenate_row(row):
            values = [str(row[col]) for col in self.levels if pd.notna(row[col]) and str(row[col]).strip()]
            return ''.join(values) if values else None

        self.df['newcol'] = self.df.apply(concatenate_row, axis=1)

    def process_data(self):
        if self.df is None or not self.column_combos:
            QMessageBox.warning(self, "خطا", "ابتدا فایل و ستون‌ها را انتخاب کنید.")
            return

        self.levels.clear()
        self.normalize_lengths()

        self.concatenate_columns()

        self.output_text.setText("نرمال‌سازی و ترکیب ستون‌ها با موفقیت انجام شد.")
        save_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", "", "Excel Files (*.xlsx)")
        if save_path:
            self.df.to_excel(save_path, index=False)
            QMessageBox.information(self, "ذخیره شد", "فایل خروجی با موفقیت ذخیره شد.")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("اکسل‌های لود حسابرسی")
        self.setGeometry(100, 100, 900, 600)

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
        self.tab_widget.addTab(CodeSplitterApp(), "تفکیک کدینگ سطوح")
        self.tab_widget.addTab(ExcelMergeApp(), " اخذ خروجی نهایی کدینگ  "  )
        self.tab_widget.addTab(CodeLengthNormalizer(), " یکسان سازی ارقام کدینگ   ")

        clean_tab = QWidget()
        clean_layout = QVBoxLayout()
        clean_layout.addWidget(DataCleanerApp())
        clean_tab.setLayout(clean_layout)
        self.tab_widget.addTab(clean_tab, "پاکسازی داده‌ها")

        format_tab = QWidget()
        format_layout = QVBoxLayout()
        format_layout.addWidget(ExcelFormatterApp())
        format_tab.setLayout(format_layout)
        self.tab_widget.addTab(format_tab, "    ااصلاح مقادیر فایل ریز اسناد (تاریخ ،مبالغ)    ")

        self.tab_widget.addTab(DiscrepancyTab(), "بررسی مغایرت‌ها")
        self.tab_widget.addTab(Level1BalanceTab(), " ساخت تراز  ")

        
        self.tab_widget.addTab(CompareExcelTab(), "مقایسه لود مرحله‌ای")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
