from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel, QFileDialog
import pandas as pd
from openpyxl import load_workbook

class Level1BalanceTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.voucher_df = None
        self.level1_df = None
        self.result_df = None

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.label_status = QLabel("وضعیت: آماده")
        self.btn_select_files = QPushButton("انتخاب فایل‌ها و شروع محاسبه تراز سطح 1")
        self.text_log = QTextEdit()
        self.text_log.setReadOnly(True)

        layout.addWidget(self.label_status)
        layout.addWidget(self.btn_select_files)
        layout.addWidget(self.text_log)

        self.setLayout(layout)

        self.btn_select_files.clicked.connect(self.run_balance_calculation)

    def log(self, msg):
        self.text_log.append(msg)

    def run_balance_calculation(self):
        self.log("انتخاب فایل‌های تراکنش...")
        voucher_files, _ = QFileDialog.getOpenFileNames(self, "انتخاب فایل‌های تراکنش‌ها (VoucherRow)", filter="Excel Files (*.xlsx *.xls)")
        if not voucher_files:
            self.log("هیچ فایلی انتخاب نشد.")
            return

        self.log("انتخاب فایل سطح 1...")
        level1_file, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل سطح 1 (AccountCoding_TS_Level1)", filter="Excel Files (*.xlsx *.xls)")
        if not level1_file:
            self.log("فایل سطح 1 انتخاب نشد.")
            return

        try:
            voucher_dfs = [pd.read_excel(f) for f in voucher_files]
            self.voucher_df = pd.concat(voucher_dfs, ignore_index=True)
            self.level1_df = pd.read_excel(level1_file)
            self.log("فایل‌ها با موفقیت بارگذاری شدند.")
        except Exception as e:
            self.log(f"خطا در خواندن فایل‌ها: {e}")
            return

        self.calculate_balance()

    def calculate_balance(self):
        if self.voucher_df is None or self.level1_df is None:
            self.log("دیتا بارگذاری نشده است.")
            return

        self.level1_df['Code'] = self.level1_df['Code'].astype(str)
        level1_length = len(self.level1_df['Code'].iloc[0])
        self.log(f"طول کد سطح 1: {level1_length}")

        self.voucher_df['Code'] = self.voucher_df['Code'].astype(str).str[:level1_length]

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
        self.log("محاسبات تراز سطح 1 انجام شد.")

        self.save_output()

    def save_output(self):
        if self.result_df is None:
            self.log("داده‌ای برای ذخیره وجود ندارد.")
            return

        output_file, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", filter="Excel Files (*.xlsx *.xls)")
        if not output_file:
            self.log("مسیر ذخیره‌سازی انتخاب نشد.")
            return

        try:
            self.result_df.to_excel(output_file, index=False)
            wb = load_workbook(output_file)
            ws = wb.active
            ws.sheet_view.rightToLeft = True
            wb.save(output_file)
            self.log(f"فایل با موفقیت ذخیره شد: {output_file}")
        except Exception as e:
            self.log(f"خطا در ذخیره‌سازی: {e}")
