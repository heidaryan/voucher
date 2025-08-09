from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel, QFileDialog
import pandas as pd
from openpyxl import load_workbook
import re
import os

class Level1BalanceTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.voucher_df = None
        self.level_df = None
        self.result_df = None
        self.level_number = "?"

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.label_status = QLabel("وضعیت: آماده")
        self.btn_select_files = QPushButton("انتخاب فایل‌ها و شروع محاسبه تراز سطح")
        self.text_log = QTextEdit()
        self.text_log.setReadOnly(True)

        layout.addWidget(self.label_status)
        layout.addWidget(self.btn_select_files)
        layout.addWidget(self.text_log)

        self.setLayout(layout)

        self.btn_select_files.clicked.connect(self.run_balance_calculation)

    def log(self, msg):
        self.text_log.append(msg)

    def extract_level_number(self, filename):
        match = re.search(r'Level(\d+)', filename, re.IGNORECASE)
        return match.group(1) if match else "?"

    def run_balance_calculation(self):
        self.log("انتخاب فایل‌های تراکنش...")
        voucher_files, _ = QFileDialog.getOpenFileNames(self, "انتخاب فایل‌های تراکنش‌ها (VoucherRow)", filter="Excel Files (*.xlsx *.xls)")
        if not voucher_files:
            self.log("هیچ فایلی انتخاب نشد.")
            return

        self.log("انتخاب فایل سطح...")
        level_file, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل (AccountCoding_TS_LevelX)", filter="Excel Files (*.xlsx *.xls)")
        if not level_file:
            self.log("فایل سطح انتخاب نشد.")
            return

        # استخراج شماره سطح از نام فایل
        self.level_number = self.extract_level_number(os.path.basename(level_file))
        self.btn_select_files.setText(f"انتخاب فایل‌ها و شروع محاسبه تراز سطح {self.level_number}")

        try:
            voucher_dfs = [pd.read_excel(f) for f in voucher_files]
            self.voucher_df = pd.concat(voucher_dfs, ignore_index=True)
            self.level_df = pd.read_excel(level_file)
            self.log("فایل‌ها با موفقیت بارگذاری شدند.")
        except Exception as e:
            self.log(f"خطا در خواندن فایل‌ها: {e}")
            return

        self.calculate_balance()

    def calculate_balance(self):
        if self.voucher_df is None or self.level_df is None:
            self.log("دیتا بارگذاری نشده است.")
            return

        self.level_df['Code'] = self.level_df['Code'].astype(str)
        level_length = len(self.level_df['Code'].iloc[0])
        self.log(f"طول کد سطح {self.level_number}: {level_length}")

        self.voucher_df['Code'] = self.voucher_df['Code'].astype(str).str[:level_length]

        grouped = self.voucher_df.groupby('Code').agg({
            'DebtorAmount': 'sum',
            'CreditorAmount': 'sum',
            'Name': 'first'
        }).reset_index()

        name_col = grouped.pop('Name')
        grouped.insert(1, 'Name', name_col)

        grouped['DebtorAmountInBeginningPeriod'] = grouped['DebtorAmount']
        grouped['CreditorAmountInBeginningPeriod'] = grouped['CreditorAmount']
        grouped['DebtorAmountInDuringPeriod'] = grouped['DebtorAmount']
        grouped['CreditorAmountInDuringPeriod'] = grouped['CreditorAmount']
        grouped['RemainingDebtorAmount'] = 0
        grouped['RemainingCreditorAmount'] = 0

        for idx, row in grouped.iterrows():
            if row['DebtorAmount'] > row['CreditorAmount']:
                grouped.at[idx, 'RemainingDebtorAmount'] = row['DebtorAmount'] - row['CreditorAmount']
            else:
                grouped.at[idx, 'RemainingCreditorAmount'] = row['CreditorAmount'] - row['DebtorAmount']

        grouped.drop(columns=['DebtorAmount', 'CreditorAmount'], inplace=True)

        self.result_df = grouped
        self.log(f"محاسبات تراز سطح {self.level_number} انجام شد.")

        self.save_output()

    def save_output(self):
        if self.result_df is None:
            self.log("داده‌ای برای ذخیره وجود ندارد.")
            return

        default_filename = f"Balance_TS_Level{self.level_number}.xlsx"
        output_file, _ = QFileDialog.getSaveFileName(self, f"ذخیره فایل خروجی تراز سطح {self.level_number}", default_filename, filter="Excel Files (*.xlsx *.xls)")
        if not output_file:
            self.log("مسیر ذخیره‌سازی انتخاب نشد.")
            return

        try:
            self.result_df.to_excel(output_file, index=False)
            wb = load_workbook(output_file)
            ws = wb.active
            ws.sheet_view.rightToLeft = True
            wb.save(output_file)
            self.log(f"فایل تراز سطح {self.level_number} با موفقیت ذخیره شد: {output_file}")
        except Exception as e:
            self.log(f"خطا در ذخیره‌سازی: {e}")
