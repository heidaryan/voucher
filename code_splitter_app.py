import pandas as pd
import xlsxwriter
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QMessageBox, QFileDialog,
    QLabel, QTableWidget, QTableWidgetItem, QSpinBox, QComboBox
)


class CodeSplitterApp(QWidget):
    def __init__(self):
        super().__init__()

        self.df = None
        self.input_file = None
        self.output_file = None

        self.layout = QVBoxLayout()

        # انتخاب فایل
        self.load_button = QPushButton("انتخاب فایل اکسل ورودی")
        self.load_button.clicked.connect(self.select_input_file)
        self.layout.addWidget(self.load_button)

        # ردیف هدر
        self.layout.addWidget(QLabel("تعیین ردیف هدر (شروع از ۰):"))
        self.header_spinbox = QSpinBox()
        self.header_spinbox.setValue(0)
        self.layout.addWidget(self.header_spinbox)

        # حد مجاز سلول خالی
        self.layout.addWidget(QLabel("حداکثر تعداد مجاز ستون‌های خالی در هر ردیف (برای حذف):"))
        self.empty_column_spinbox = QSpinBox()
        self.empty_column_spinbox.setValue(10)
        self.layout.addWidget(self.empty_column_spinbox)

        # دکمه بارگذاری
        self.preview_button = QPushButton("نمایش جدول با هدر انتخابی")
        self.preview_button.clicked.connect(self.load_with_selected_header)
        self.layout.addWidget(self.preview_button)

        # انتخاب ستون
        self.layout.addWidget(QLabel("انتخاب ستونی برای پردازش کد:"))
        self.column_combobox = QComboBox()
        self.layout.addWidget(self.column_combobox)

        # کاراکتر حذف‌شونده
        self.layout.addWidget(QLabel("کاراکتری که باید حذف شود (قابل انتخاب یا تایپ):"))
        self.remove_char_combobox = QComboBox()
        self.remove_char_combobox.setEditable(True)
        self.remove_char_combobox.addItems(['/', '-', '.', '_'])
        self.layout.addWidget(self.remove_char_combobox)

        # تعداد رقم جداشونده
        self.layout.addWidget(QLabel("تعداد رقم اول که جدا شود:"))
        self.split_length_spinbox = QSpinBox()
        self.split_length_spinbox.setValue(3)
        self.layout.addWidget(self.split_length_spinbox)

        # جدول پیش‌نمایش
        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        # دکمه پردازش و ذخیره
        self.process_button = QPushButton("شروع پردازش و تقسیم کد")
        self.process_button.clicked.connect(self.process_and_save)
        self.layout.addWidget(self.process_button)

        self.setLayout(self.layout)

    def show_message(self, title, message):
        QMessageBox.information(self, title, message)

    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل ورودی", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.input_file = file_path
            self.show_message("موفقیت", f"فایل انتخاب شد:\n{self.input_file}")

    def load_with_selected_header(self):
        if not self.input_file:
            self.show_message("خطا", "ابتدا فایل اکسل ورودی را انتخاب کنید.")
            return

        header_row = self.header_spinbox.value()
        max_empty_allowed = self.empty_column_spinbox.value()

        try:
            df = pd.read_excel(self.input_file, header=header_row, dtype=str)
            df_cleaned = df[df.isnull().sum(axis=1) <= max_empty_allowed].copy()
            self.df = df_cleaned

            self.update_column_combobox()
            self.show_table(self.df)

            removed_rows = len(df) - len(self.df)
            if removed_rows > 0:
                self.show_message("پاکسازی انجام شد", f"{removed_rows} ردیف به‌دلیل داشتن بیش از {max_empty_allowed} مقدار خالی حذف شدند.")
        except Exception as e:
            self.show_message("خطا", f"خطا در بارگذاری فایل:\n{e}")

    def update_column_combobox(self):
        self.column_combobox.clear()
        if self.df is not None:
            self.column_combobox.addItems(self.df.columns.astype(str))

    def show_table(self, df):
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.astype(str))

        for i, row in df.head(5).iterrows():
            self.table.insertRow(i)
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(i, j, item)

    def process_and_save(self):
        if self.df is None:
            self.show_message("خطا", "ابتدا فایل را بارگذاری کنید.")
            return

        selected_column = self.column_combobox.currentText()
        if selected_column not in self.df.columns:
            self.show_message("خطا", f"ستون '{selected_column}' پیدا نشد.")
            return

        # حذف ردیف‌هایی که در ستون انتخابی مقدار خالی یا فقط فاصله دارند
        self.df = self.df[self.df[selected_column].notna() & (self.df[selected_column].astype(str).str.strip() != '')].copy()

        char_to_remove = self.remove_char_combobox.currentText()
        split_length = self.split_length_spinbox.value()

        try:
            # حذف کاراکتر و جداسازی کد
            self.df[selected_column] = self.df[selected_column].astype(str).str.replace(char_to_remove, '', regex=False)
            self.df['code_part1'] = self.df[selected_column].str[:split_length]
            self.df['code_part2'] = self.df[selected_column].str[split_length:]
        except Exception as e:
            self.show_message("خطا", f"خطا در پردازش ستون:\n{e}")
            return

        output_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", "", "Excel Files (*.xlsx)")
        if not output_path:
            self.show_message("خطا", "مسیر ذخیره انتخاب نشد.")
            return

        try:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                self.df.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                text_format = workbook.add_format({'num_format': '@'})

                for col_num, _ in enumerate(self.df.columns.values):
                    worksheet.set_column(col_num, col_num, 20, text_format)

            self.show_message("موفقیت", f"فایل با موفقیت ذخیره شد:\n{output_path}")
        except Exception as e:
            self.show_message("خطا", f"خطا در ذخیره فایل:\n{e}")
