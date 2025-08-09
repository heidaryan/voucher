import pandas as pd
import random
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QComboBox,
    QMessageBox, QSpinBox
)


class CodeLengthNormalizer(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        self.levels = {}
        self.combo_boxes = []

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.label = QLabel("ابتدا یک فایل اکسل را انتخاب کنید:")
        self.layout.addWidget(self.label)

        self.load_button = QPushButton("انتخاب فایل اکسل")
        self.load_button.clicked.connect(self.load_excel)
        self.layout.addWidget(self.load_button)

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در خواندن فایل:\n{e}")
            return

        self.build_column_selection_ui()

    def build_column_selection_ui(self):
        self.clear_layout()

        self.layout.addWidget(QLabel("تعداد ستون‌هایی که می‌خواهید انتخاب کنید (۱ تا ۴):"))
        self.spin_box = QSpinBox()
        self.spin_box.setRange(1, 4)
        self.layout.addWidget(self.spin_box)

        self.generate_combo_button = QPushButton("نمایش ستون‌ها")
        self.generate_combo_button.clicked.connect(self.create_combo_boxes)
        self.layout.addWidget(self.generate_combo_button)

    def create_combo_boxes(self):
        self.levels.clear()
        self.combo_boxes.clear()

        for i in range(self.spin_box.value()):
            combo = QComboBox()
            combo.addItems(["- انتخاب کنید -"] + list(self.df.columns))
            self.combo_boxes.append(combo)
            self.layout.addWidget(combo)

        self.additional_column_combo = QComboBox()
        self.additional_column_combo.addItems(["- انتخاب کنید -"] + list(self.df.columns))
        self.layout.addWidget(QLabel("ستون اضافی برای مقایسه را انتخاب کنید:"))
        self.layout.addWidget(self.additional_column_combo)

        self.ok_button = QPushButton("انجام نرمال‌سازی")
        self.ok_button.clicked.connect(self.perform_normalization)
        self.layout.addWidget(self.ok_button)

    def perform_normalization(self):
        if self.df is None:
            QMessageBox.warning(self, "خطا", "فایل اکسل بارگذاری نشده است.")
            return

        self.levels.clear()
        for combo in self.combo_boxes:
            col = combo.currentText()
            if col and col != "- انتخاب کنید -":
                max_length = self.df[col].astype(str).str.len().max()
                self.levels[col] = max_length

        self.additional_column = self.additional_column_combo.currentText()
        if self.additional_column == "- انتخاب کنید -" or not self.additional_column:
            QMessageBox.warning(self, "خطا", "باید یک ستون اضافی انتخاب شود.")
            return

        if len(self.levels) < 1:
            QMessageBox.warning(self, "خطا", "حداقل یک ستون باید انتخاب شود.")
            return

        self.normalize_lengths()
        self.compare_rows()
        self.concatenate_columns()

        save_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", "", "Excel Files (*.xlsx)")
        if save_path:
            self.df.to_excel(save_path, index=False)
            QMessageBox.information(self, "موفق", "فایل خروجی با موفقیت ذخیره شد.")

    def normalize_lengths(self):
        for col, max_length in self.levels.items():
            self.df[col] = self.df[col].apply(
                lambda x: str(int(x)).strip().zfill(max_length) if pd.notna(x) and str(x).strip() != '' else ''
            )

    def generate_value_based_on_length(self, value, row_index):
        random.seed(row_index)
        length = len(value)
        return str(random.randint(10**(length-1), 10**length - 1)) if length > 0 else '1'

    def concatenate_columns(self):
        def concatenate_row(row):
            values = [str(row[col]) for col in self.levels if pd.notna(row[col]) and str(row[col]).strip() != '']
            return ''.join(values) if values else None

        self.df['newcol'] = self.df.apply(concatenate_row, axis=1)

    def compare_rows(self):
        last_column = list(self.levels.keys())[-1]
        used_values = {}

        compare_columns = list(self.levels.keys()) + [self.additional_column]

        for i, row in self.df.iterrows():
            last_val = str(row[last_column]).strip()
            if last_val == '0' * len(last_val):
                valid = any(pd.notna(row[col]) and str(row[col]).strip() != '' for col in compare_columns)
                if not valid:
                    continue

                mask = self.df[compare_columns].astype(str).applymap(str.strip).eq(row[compare_columns].astype(str).map(str.strip)).all(axis=1)
                similar = self.df[mask]

                if len(similar) > 0:
                    if i not in used_values:
                        used_values[i] = self.generate_value_based_on_length(row[last_column], i)

                    for idx in similar.index:
                        if self.df.at[idx, last_column] == '0' * len(self.df.at[idx, last_column]):
                            self.df.at[idx, last_column] = used_values[i]

    def clear_layout(self):
        for i in reversed(range(self.layout.count())):
            widget = self.layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
