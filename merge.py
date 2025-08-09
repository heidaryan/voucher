import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QLabel, QComboBox, QTextEdit, QMessageBox
)

class ExcelMergeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ادغام دو اکسل با گزارش کدهای گمشده")
        self.resize(700, 600)

        self.df1 = None  # اکسل اول
        self.df2 = None  # اکسل دوم
        self.df2_path = None  # مسیر فایل دوم

        layout = QVBoxLayout()

        # بارگذاری و انتخاب ستون اکسل اول
        self.btn_load_1 = QPushButton("بارگذاری فایل ریز اسناد")
        self.btn_load_1.clicked.connect(self.load_file1)
        layout.addWidget(self.btn_load_1)

        layout.addWidget(QLabel("انتخاب ستون کد در اکسل ریزاسناد:"))
        self.combo_code_1 = QComboBox()
        layout.addWidget(self.combo_code_1)

        # بارگذاری و انتخاب ستون‌های اکسل دوم
        self.btn_load_2 = QPushButton("بارگذاری فایل تراز(کل/معین/تراز)")
        self.btn_load_2.clicked.connect(self.load_file2)
        layout.addWidget(self.btn_load_2)

        # پیش‌نمایش فایل دوم
        layout.addWidget(QLabel("انتخاب ستون Code در تراز"))
        self.text_preview_2 = QTextEdit()
        self.text_preview_2.setReadOnly(True)
        layout.addWidget(self.text_preview_2)

        layout.addWidget(QLabel("انتخاب ردیف هدر در اکسل تراز:"))
        self.combo_header_row = QComboBox()
        self.combo_header_row.addItems([str(i) for i in range(20)])  # از 0 تا 9
        self.combo_header_row.currentIndexChanged.connect(self.reload_df2_with_header)
        layout.addWidget(self.combo_header_row)

        layout.addWidget(QLabel("انتخاب ستون Code در اکسل تراز:"))
        self.combo_code_2 = QComboBox()
        layout.addWidget(self.combo_code_2)

        layout.addWidget(QLabel("انتخاب ستون Name در اکسل تراز:"))
        self.combo_name_2 = QComboBox()
        layout.addWidget(self.combo_name_2)

        # دکمه مرج و گزارش
        self.btn_merge = QPushButton("ادغام و گزارش کدهای گمشده")
        self.btn_merge.clicked.connect(self.merge_and_report)
        layout.addWidget(self.btn_merge)

        # محل نمایش گزارش
        layout.addWidget(QLabel("گزارش کدهای گمشده:"))
        self.text_report = QTextEdit()
        self.text_report.setReadOnly(True)
        layout.addWidget(self.text_report)

        self.setLayout(layout)

    def load_file1(self):
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل اول", filter="Excel Files (*.xlsx *.xls)")
        if path:
            try:
                self.df1 = pd.read_excel(path, dtype=str)
                self.combo_code_1.clear()
                self.combo_code_1.addItems(self.df1.columns.astype(str))
                QMessageBox.information(self, "موفقیت", "فایل اول بارگذاری شد.")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در بارگذاری فایل اول:\n{e}")

    def load_file2(self):
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل دوم", filter="Excel Files (*.xlsx *.xls)")
        if path:
            try:
                self.df2_path = path
                preview_df = pd.read_excel(path, header=None, dtype=str, nrows=10)
                self.text_preview_2.setPlainText(preview_df.to_string(index=False))
                self.reload_df2_with_header()
                QMessageBox.information(self, "موفقیت", "پیش‌نمایش فایل دوم بارگذاری شد.")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در بارگذاری فایل دوم:\n{e}")

    def reload_df2_with_header(self):
        if self.df2_path:
            try:
                header_row = int(self.combo_header_row.currentText())
                self.df2 = pd.read_excel(self.df2_path, header=header_row, dtype=str)

                self.combo_code_2.clear()
                self.combo_name_2.clear()
                self.combo_code_2.addItems(self.df2.columns.astype(str))
                self.combo_name_2.addItems(self.df2.columns.astype(str))
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در بارگذاری با هدر انتخابی:\n{e}")

    def merge_and_report(self):
        if self.df1 is None or self.df2 is None:
            QMessageBox.warning(self, "هشدار", "لطفاً هر دو فایل را بارگذاری کنید.")
            return

        code_col_1 = self.combo_code_1.currentText()
        code_col_2 = self.combo_code_2.currentText()
        name_col_2 = self.combo_name_2.currentText()

        if not code_col_1 or not code_col_2 or not name_col_2:
            QMessageBox.warning(self, "هشدار", "لطفاً ستون‌ها را انتخاب کنید.")
            return

        try:
            # مرج بر اساس ستون کد
            merged = pd.merge(
                self.df1,
                self.df2[[code_col_2, name_col_2]],
                left_on=code_col_1,
                right_on=code_col_2,
                how='left'
            )

            merged.rename(columns={name_col_2: "نام از اکسل دوم"}, inplace=True)

            missing_codes = merged[merged["نام از اکسل دوم"].isna()][code_col_1].unique()

            report = ""
            if len(missing_codes) > 0:
                report += f"کدهایی که در اکسل دوم یافت نشدند ({len(missing_codes)} مورد):\n"
                for c in missing_codes:
                    report += f" - {c}\n"
            else:
                report = "تمام کدهای اکسل اول در اکسل دوم موجود هستند."

            self.text_report.setPlainText(report)

            save_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", filter="Excel Files (*.xlsx)")
            if save_path:
                merged.to_excel(save_path, index=False)
                QMessageBox.information(self, "موفقیت", f"فایل با موفقیت ذخیره شد:\n{save_path}")

        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در مرج و گزارش:\n{e}")


