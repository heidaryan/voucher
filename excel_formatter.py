# excel_formatter.py
import pandas as pd
import jdatetime
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QMessageBox, QFileDialog

class ExcelFormatterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self.input_file = None
        self.output_file = None

        layout = QVBoxLayout()

        self.load_button = QPushButton("انتخاب فایل ورودی اکسل")
        self.load_button.clicked.connect(self.select_input_file)
        layout.addWidget(self.load_button)

        self.save_button = QPushButton("انتخاب مسیر ذخیره فایل خروجی")
        self.save_button.clicked.connect(self.select_output_file)
        layout.addWidget(self.save_button)

        self.process_button = QPushButton("شروع پردازش")
        self.process_button.clicked.connect(self.process_data_and_save)
        layout.addWidget(self.process_button)

        self.setLayout(layout)

    def show_message(self, title, message):
        QMessageBox.information(self, title, message)

    def select_input_file(self):
        self.input_file, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل ورودی", "", "Excel Files (*.xlsx *.xls)")
        if self.input_file:
            self.df = pd.read_excel(self.input_file, dtype={'CreditorAmount': float, 'DebtorAmount': float})
            self.show_message("موفقیت", f"فایل ورودی بارگذاری شد:\n{self.input_file}")

    def select_output_file(self):
        self.output_file, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", "", "Excel Files (*.xlsx)")
        if self.output_file:
            self.show_message("موفقیت", f"فایل خروجی انتخاب شد:\n{self.output_file}")

    def format_numeric_date(self, date_str):
        try:
            if pd.isna(date_str):
                return date_str
            if isinstance(date_str, (int, float)) or (isinstance(date_str, str) and date_str.isdigit() and len(date_str) == 8):
                year = str(date_str)[:4]
                month = str(date_str)[4:6]
                day = str(date_str)[6:8]
                return f"{year}/{month}/{day}"
            return date_str
        except Exception as e:
            print(f"خطا در تبدیل تاریخ: {e}")
            return date_str

    def convert_to_farsi_date(self, date_str):
        try:
            if pd.isna(date_str):
                return date_str
            year = int(date_str[:4])
            month = int(date_str[4:6])
            day = int(date_str[6:8])
            gregorian_date = jdatetime.date(year, month, day)
            return gregorian_date.strftime('%Y/%m/%d')
        except Exception as e:
            print(f"خطا در تبدیل تاریخ میلادی به شمسی: {e}")
            return date_str

    def convert_dates(self, df, date_columns):
        for column in date_columns:
            if column in df.columns:
                df[column] = df[column].apply(lambda x: self.format_numeric_date(x))

    def process_data_and_save(self):
        if self.df is not None:
            self.convert_dates(self.df, ['PersianVoucherDate'])
            self.df['PersianVoucherDate'] = self.df['PersianVoucherDate'].apply(lambda x: self.convert_to_farsi_date(x))

            if 'CreditorAmount' in self.df.columns and 'DebtorAmount' in self.df.columns:
                self.df['CreditorAmount'] = self.df['CreditorAmount'].apply(lambda x: '{:.0f}'.format(x) if not pd.isna(x) else x)
                self.df['DebtorAmount'] = self.df['DebtorAmount'].apply(lambda x: '{:.0f}'.format(x) if not pd.isna(x) else x)
            else:
                self.show_message("خطا", "ستون‌های مورد نیاز وجود ندارند.")

            if self.output_file:
                self.df.to_excel(self.output_file, index=False, float_format='%.0f')
                self.show_message("موفقیت", f"فایل اصلاح‌شده ذخیره شد:\n{self.output_file}")
            else:
                self.show_message("خطا", "مسیر فایل خروجی انتخاب نشده.")
        else:
            self.show_message("خطا", "فایل ورودی بارگذاری نشده است.")
