import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTableWidget, QTableWidgetItem,
    QPushButton, QVBoxLayout, QWidget, QLabel, QSpinBox, QListWidget, QMessageBox,
    QInputDialog, QComboBox , QDialog , QDialogButtonBox,QScrollArea
)
class DataCleanerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("پاکسازی داده‌ها")
        self.resize(800, 600)

        # ویجت اصلی
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # لایه‌ها
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        # جدول برای نمایش داده‌ها
        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        # کنترل‌های انتخاب فایل
        self.file_label = QLabel("انتخاب فایل اکسل برای پردازش:")
        self.layout.addWidget(self.file_label)

        self.load_button = QPushButton("بارگذاری فایل اکسل")
        self.load_button.clicked.connect(self.load_excel_file)
        self.layout.addWidget(self.load_button)

        # انتخاب ردیف هدر
        self.header_label = QLabel("تعیین ردیف هدر (پیش‌فرض: ۰):")
        self.layout.addWidget(self.header_label)

        self.header_spinbox = QSpinBox()
        self.header_spinbox.setValue(0)
        self.layout.addWidget(self.header_spinbox)

        # حد آستانه داده‌های گمشده
        self.missing_label = QLabel("حذف ردیف‌هایی با بیش از تعداد مشخص داده‌های گمشده (پیش‌فرض: ۴):")
        self.layout.addWidget(self.missing_label)

        self.missing_spinbox = QSpinBox()
        self.missing_spinbox.setValue(4)
        self.layout.addWidget(self.missing_spinbox)

        # دکمه‌های پردازش و ذخیره‌سازی
        self.process_button = QPushButton("پردازش داده‌ها")
        self.process_button.clicked.connect(self.process_data)
        self.layout.addWidget(self.process_button)

        # دکمه به‌روزرسانی هدرها
        self.update_headers_button = QPushButton("به‌روزرسانی هدرها")
        self.update_headers_button.clicked.connect(self.update_headers)
        self.layout.addWidget(self.update_headers_button)


        self.column_list = QListWidget()
        self.column_list.setSelectionMode(QListWidget.MultiSelection)
        self.layout.addWidget(self.column_list)

        self.rename_headers_button = QPushButton("تغییر نام هدرها")
        self.rename_headers_button.clicked.connect(self.rename_headers)
        self.layout.addWidget(self.rename_headers_button)

        self.finalize_button = QPushButton("تشکیل اکسل نهایی")
        self.finalize_button.clicked.connect(self.process_columns)
        self.layout.addWidget(self.finalize_button)


        self.finall_button = QPushButton("ساخت فلگ")
        self.finall_button.clicked.connect(self.apply_voucher_flag)
        self.layout.addWidget(self.finall_button)



        self.fill_row_button = QPushButton("پر کردن ستون RowNumber")
        self.fill_row_button.clicked.connect(self.fill_row_number_with_ones)
        self.layout.addWidget(self.fill_row_button)


        self.clos_button = QPushButton("حذف حساب‌های بسته شده")
        self.clos_button.clicked.connect(self.remove_rows_based_on_description)
        self.layout.addWidget(self.clos_button)


        self.save_button = QPushButton("ذخیره فایل پردازش‌شده")
        self.save_button.clicked.connect(self.save_file)
        self.save_button.setEnabled(False)
        self.layout.addWidget(self.save_button)

        # متغیرهای داده
        self.dataframe = None
        self.cleaned_dataframe = None

        self.standard_columns = [
            'Code', 'Name', 'DebtorAmount', 'CreditorAmount', 'VoucherNumber', 'RowNumber', 
            'PersianVoucherDate', 'DescriptionRow', 'VoucherType_Flag', 
            'Matrix_1_Code', 'Matrix_1_Name', 'Matrix_2_Code', 'Matrix_2_Name', 
            'Matrix_3_Code', 'Matrix_3_Name', 'Matrix_4_Code', 'Matrix_4_Name', 
            'Matrix_5_Code', 'Matrix_5_Name', 'Matrix_6_Code', 'Matrix_6_Name'
        ]

    def load_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "باز کردن فایل اکسل", "", "فایل‌های اکسل (*.xlsx *.xls)")
        if file_path:
            try:
                self.dataframe = pd.read_excel(file_path, header=None)
                self.populate_table(self.dataframe)
                self.populate_column_list()
                self.save_button.setEnabled(False)
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در بارگذاری فایل: {e}")


    def populate_table(self, dataframe):

        if dataframe is not None:
            # فقط 5 ردیف اول را انتخاب می‌کنیم
            dataframe_head = dataframe.head(5)

            # تنظیم تعداد ردیف‌ها و ستون‌ها براساس تعداد ردیف‌ها و ستون‌های 5 ردیف اول
            self.table.setRowCount(dataframe_head.shape[0])
            self.table.setColumnCount(dataframe_head.shape[1])
            self.table.setHorizontalHeaderLabels([str(i) for i in range(dataframe_head.shape[1])])

            # پر کردن سلول‌ها با داده‌ها
            for row in range(dataframe_head.shape[0]):
                for col in range(dataframe_head.shape[1]):
                    value = dataframe_head.iloc[row, col]
                    self.table.setItem(row, col, QTableWidgetItem(str(value) if not pd.isnull(value) else ""))
        else:
            QMessageBox.warning(None, "خطا", "دیتافریم موجود نیست!")

    def update_headers(self):
        if self.dataframe is not None:
            header_row = self.header_spinbox.value()
            headers = self.dataframe.iloc[header_row].tolist()
            self.updated_dataframe = self.dataframe[header_row + 1:].copy()
            self.updated_dataframe.columns = headers
            self.populate_table(self.updated_dataframe)
            self.column_list.clear()
            self.column_list.addItems(headers)
            QMessageBox.information(self, "موفقیت", "هدرها با موفقیت به‌روزرسانی شدند!")
        else:
            QMessageBox.warning(self, "خطا", "ابتدا یک فایل بارگذاری کنید!")

    def populate_column_list(self):
        if self.dataframe is not None:
            self.column_list.clear()
            header_row = self.header_spinbox.value()
            headers = self.dataframe.iloc[header_row].tolist()
            self.column_list.addItems([str(header) for header in headers])

    def process_data(self):
        if self.dataframe is not None:
            try:
                header_row = self.header_spinbox.value()
                headers = self.dataframe.iloc[header_row].tolist()

                df_cleaned = self.dataframe[~self.dataframe.apply(lambda row: row.tolist() == headers, axis=1)]
                df_cleaned.columns = headers
                df_cleaned = df_cleaned.iloc[header_row :]

                missing_threshold = self.missing_spinbox.value()
                df_cleaned = df_cleaned[df_cleaned.isnull().sum(axis=1) <= missing_threshold]

                selected_columns = [item.text() for item in self.column_list.selectedItems()]
                if selected_columns:
                    df_cleaned = df_cleaned.drop(columns=selected_columns, errors='ignore')

                self.cleaned_dataframe = df_cleaned
                self.populate_table(self.cleaned_dataframe)
                self.save_button.setEnabled(True)
                QMessageBox.information(self, "موفقیت", "داده‌ها با موفقیت پردازش شدند!")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در پردازش داده‌ها: {e}")


    def rename_headers(self):
        if self.cleaned_dataframe is not None:
            try:
                current_headers = self.cleaned_dataframe.columns.tolist()
                dialog = QDialog(self)
                dialog.setWindowTitle("تغییر نام ستون‌ها")
                layout = QVBoxLayout()
                new_headers = {}

                for header in current_headers:
                    label = QLabel(f"نام جدید برای ستون '{header}':")
                    combo_box = QComboBox()
                    combo_box.addItem(header)  # افزودن نام فعلی به عنوان گزینه پیش‌فرض
                    combo_box.addItem("Code")
                    combo_box.addItem("Name")
                    combo_box.addItem("DebtorAmount")
                    combo_box.addItem("CreditorAmount")
                    combo_box.addItem("VoucherNumber")
                    combo_box.addItem("RowNumber")
                    combo_box.addItem("PersianVoucherDate")
                    combo_box.addItem("DescriptionRow")
                    combo_box.addItem("VoucherType_Flag")
                    combo_box.addItem("Matrix_1_Code")
                    combo_box.addItem("Matrix_1_Name")
                    combo_box.addItem("Matrix_2_Code")
                    combo_box.addItem("Matrix_2_Name")
                    combo_box.addItem("Matrix_3_Code")
                    combo_box.addItem("Matrix_3_Name")
                    combo_box.addItem("Matrix_4_Code")
                    combo_box.addItem("Matrix_4_Name")
                    combo_box.addItem("Matrix_5_Code")
                    combo_box.addItem("Matrix_5_Name")
                    combo_box.addItem("Matrix_6_Code")
                    combo_box.addItem("Matrix_6_Name")

                    new_headers[header] = combo_box
                    layout.addWidget(label)
                    layout.addWidget(combo_box)

                dialog.setLayout(layout)
                button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                layout.addWidget(button_box)

                button_box.accepted.connect(dialog.accept)
                button_box.rejected.connect(dialog.reject)

                if dialog.exec_() == QDialog.Accepted:
                    new_column_names = {header: new_headers[header].currentText() for header in current_headers}
                    self.cleaned_dataframe.rename(columns=new_column_names, inplace=True)
                    self.populate_table(self.cleaned_dataframe)

                    QMessageBox.information(self, "موفقیت", "هدرها با موفقیت تغییر یافتند!")
                else:
                    QMessageBox.warning(self, "عملیات لغو شد", "تغییر نام هدرها لغو شد.")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در تغییر نام هدرها: {e}")



    def process_columns(self):
        if self.cleaned_dataframe is not None:
            try:
                # حذف ستون‌های اضافی
                current_columns = set(self.cleaned_dataframe.columns)
                standard_columns_set = set(self.standard_columns)
                extra_columns = current_columns - standard_columns_set
                if extra_columns:
                    self.cleaned_dataframe = self.cleaned_dataframe.drop(columns=extra_columns)

                # اضافه کردن ستون‌های گم‌شده
                for col in self.standard_columns:
                    if col not in self.cleaned_dataframe.columns:
                        self.cleaned_dataframe[col] = None

                # تنظیم ترتیب ستون‌ها مطابق با نمونه
                self.cleaned_dataframe = self.cleaned_dataframe[self.standard_columns]

                # به‌روزرسانی جدول (اینجا متدی مثل populate_table فرض شده)
                self.populate_table(self.cleaned_dataframe)
                QMessageBox.information(None, "موفقیت", "ستون‌های اضافی حذف شدند، ستون‌های گم‌شده اضافه شدند و ترتیب تنظیم شد!")
            except Exception as e:
                QMessageBox.critical(None, "خطا", f"خطا در پردازش ستون‌ها: {e}")
        else:
            QMessageBox.warning(None, "خطا", "دیتافریم پردازش‌شده‌ای وجود ندارد!")



    def apply_voucher_flag(self):
        if self.cleaned_dataframe is not None:
            try:
                # بررسی وجود ستون‌های مورد نیاز
                if 'DescriptionRow' in self.cleaned_dataframe.columns and 'VoucherType_Flag' in self.cleaned_dataframe.columns:
                    keywords = ['افتتاح', 'افتاح', 'افتتاحیه', 'افتاحیه', 'افتتاحییه', 'افتاحییه']
                    
                    # اعمال تغییرات روی ستون VoucherType_Flag
                    self.cleaned_dataframe['VoucherType_Flag'] = self.cleaned_dataframe['DescriptionRow'].apply(
                        lambda x: 2 if any(keyword in str(x) for keyword in keywords) else 1
                    )

                    # پیام موفقیت
                    QMessageBox.information(None, "موفقیت", "ستون VoucherType_Flag با موفقیت به‌روزرسانی شد!")
                else:
                    QMessageBox.warning(None, "خطا", "ستون‌های DescriptionRow و VoucherType_Flag وجود ندارند!")
            except Exception as e:
                QMessageBox.critical(None, "خطا", f"خطا در اعمال تغییرات روی ستون‌ها: {e}")
        else:
            QMessageBox.warning(None, "خطا", "دیتافریم پردازش‌شده‌ای وجود ندارد!")



    def fill_row_number_with_ones(self):
        if self.cleaned_dataframe is not None:
            try:
                # پر کردن ستون RowNumber با مقدار 1
                if 'RowNumber' in self.cleaned_dataframe.columns:
                    self.cleaned_dataframe['RowNumber'] = 1
                    self.populate_table(self.cleaned_dataframe)
                    QMessageBox.information(self, "موفقیت", "ستون RowNumber با موفقیت پر شد!")
                else:
                    QMessageBox.warning(self, "خطا", "ستون RowNumber در دیتافریم موجود نیست!")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در پر کردن ستون RowNumber: {e}")
        else:
            QMessageBox.warning(self, "خطا", "دیتافریم پردازش‌شده‌ای وجود ندارد!")



    def remove_rows_based_on_description(self):
        """
        این متد ردیف‌هایی که مقدار DescriptionRow برابر با description_value باشد را حذف می‌کند.
        """
        if self.cleaned_dataframe is not None:
            try:
                # گرفتن ورودی از کاربر
                description_value, ok = QInputDialog.getText(None, "ورود مقدار", "مقدار مورد نظر برای حذف را وارد کنید:")

                if ok and description_value:  # اگر کاربر تایید کرد و مقداری وارد کرد
                    # شناسایی ردیف‌هایی که قرار است حذف شوند
                    rows_to_delete = self.cleaned_dataframe[self.cleaned_dataframe['DescriptionRow'] == description_value]

                    # اگر ردیفی برای حذف وجود داشته باشد
                    if not rows_to_delete.empty:
                        # باز کردن دیالوگ برای انتخاب محل ذخیره فایل
                        save_path, _ = QFileDialog.getSaveFileName(None, "ذخیره فایل حذف‌شده", "", "Excel Files (*.xlsx)")

                        if save_path:  # اگر کاربر مسیری را انتخاب کرده باشد
                            # ذخیره ردیف‌های حذف‌شده در فایل اکسل با نام ثابت
                            rows_to_delete.to_excel(save_path, index=False)

                            # حذف ردیف‌هایی که در ستون DescriptionRow مقدار description_value دارند
                            self.cleaned_dataframe = self.cleaned_dataframe[self.cleaned_dataframe['DescriptionRow'] != description_value]
                            
                            # پیام موفقیت
                            QMessageBox.information(None, "موفقیت", "ردیف‌ها با موفقیت حذف شدند و در فایل ذخیره شدند!")
                        else:
                            QMessageBox.warning(None, "خطا", "مسیر ذخیره فایل انتخاب نشد.")
                    else:
                        QMessageBox.warning(None, "هشدار", "هیچ ردیفی با مقدار وارد شده پیدا نشد.")
                else:
                    QMessageBox.warning(None, "خطا", "مقدار ورودی معتبر نیست یا کاربر آن را لغو کرده است!")
            except Exception as e:
                QMessageBox.critical(None, "خطا", f"خطا در حذف ردیف‌ها: {e}")
        else:
            QMessageBox.warning(None, "خطا", "دیتافریم پردازش‌شده‌ای وجود ندارد!")



    def save_file(self):
        if self.cleaned_dataframe is not None:
            file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل اکسل", "", "فایل‌های اکسل (*.xlsx)")
            if file_path:
                try:
                    self.cleaned_dataframe.to_excel(file_path, index=False)
                    QMessageBox.information(self, "موفقیت", "فایل با موفقیت ذخیره شد!")
                except Exception as e:
                    QMessageBox.critical(self, "خطا", f"خطا در ذخیره فایل: {e}")
        else:
            QMessageBox.warning(self, "خطا", "لطفاً ابتدا داده‌ها را پردازش کنید!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataCleanerApp()
    window.show()
    sys.exit(app.exec_())
