import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client as win32
import os
import locale
from datetime import datetime
from PIL import Image, ImageTk
import pandas as pd


class AISToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("АИС ОЗП")
        self.root.geometry("1200x800")

        # Переменные
        self.selected_address = None
        self.passport_path = ""
        self.inspection_path = ""
        self.photos = []
        self.current_photo_index = 0
        self.excel_app = None
        self.passport_wb = None

        # Локаль для русской даты
        try:
            locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
        except:
            try:
                locale.setlocale(locale.LC_TIME, 'Russian_Russia.1251')
            except:
                pass

        # Интерфейс
        self.create_gui()

    def create_gui(self):
        # Левый фрейм — список домов + организация
        left_frame = ttk.Frame(self.root, width=300, height=800, padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Выбор дома
        ttk.Label(left_frame, text="Выберите адрес дома:", font=("Arial", 12)).pack(anchor=tk.W)
        self.listbox = tk.Listbox(left_frame, width=40, height=25)
        self.listbox.pack(pady=10)

        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.configure(yscrollcommand=scrollbar.set)

        # Радиокнопки организаций
        org_frame = ttk.LabelFrame(left_frame, text="Организация", padding=10)
        org_frame.pack(pady=10, fill=tk.X)

        self.org_var = tk.StringVar(value="Одинцовское")
        ttk.Radiobutton(org_frame, text="Одинцовское", variable=self.org_var, value="Одинцовское").pack(anchor=tk.W)
        ttk.Radiobutton(org_frame, text="Барвиха", variable=self.org_var, value="Барвиха").pack(anchor=tk.W)
        ttk.Radiobutton(org_frame, text="Жаворонки", variable=self.org_var, value="Жаворонки").pack(anchor=tk.W)

        # Кнопки действий
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Открыть паспорт", width=20, command=self.open_passport).pack(pady=5)
        ttk.Button(button_frame, text="Заполнить данные", width=20, command=self.fill_data_from_inspection).pack(pady=5)
        ttk.Button(button_frame, text="Показать фото", width=20, command=self.show_photos).pack(pady=5)
        ttk.Button(button_frame, text="Сохранить и закрыть", width=20, command=self.save_and_close_passport).pack(pady=5)

        # Правый фрейм — форма
        right_frame = ttk.Frame(self.root, padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Дата заполнения
        date_frame = ttk.LabelFrame(right_frame, text="Дата заполнения", padding=10)
        date_frame.pack(fill=tk.X, pady=5)
        self.date_entry = ttk.Entry(date_frame, width=30)
        self.date_entry.insert(0, datetime.now().strftime("%d %B %Y"))
        self.date_entry.pack()

        # Поставщики
        supplier_frame = ttk.LabelFrame(right_frame, text="Поставщики", padding=10)
        supplier_frame.pack(fill=tk.X, pady=5)
        ttk.Label(supplier_frame, text="Теплоснабжение:").pack(anchor=tk.W)
        self.supplier1 = ttk.Entry(supplier_frame, width=50)
        self.supplier1.insert(0, "АО \"Мосэнергосбыт\"")
        self.supplier1.pack()
        ttk.Label(supplier_frame, text="Энергоснабжение:").pack(anchor=tk.W)
        self.supplier2 = ttk.Entry(supplier_frame, width=50)
        self.supplier2.insert(0, "ООО \"Пример организация\"")
        self.supplier2.pack()

        # Чекбоксы — объемы работ
        options_frame = ttk.LabelFrame(right_frame, text="Объемы работ", padding=10)
        options_frame.pack(fill=tk.X, pady=5)

        self.option_vars = {}
        for option in [
            "Теплоснабжение",
            "Холодное водоснабжение",
            "Энергоснабжение",
            "Договор ВДГО",
            "Собственник",
            "Промывка",
            "Промывка ГВС",
            "Промывка ХВС",
            "Промывка КС"
        ]:
            var = tk.BooleanVar()
            self.option_vars[option] = var
            ttk.Checkbutton(options_frame, text=option, variable=var).pack(anchor=tk.W)

        # Текстовые поля
        text_fields_frame = ttk.Frame(right_frame)
        text_fields_frame.pack(fill=tk.X, pady=10)

        tk.Label(text_fields_frame, text="Центральное ОАО 'Одинцовская теплосеть'").pack(anchor=tk.W)
        self.txb_eto = ttk.Entry(text_fields_frame, width=50)
        self.txb_eto.pack(pady=5)

        tk.Label(text_fields_frame, text="Текущая дата").pack(anchor=tk.W)
        self.txb_date = ttk.Entry(text_fields_frame, width=50)
        self.txb_date.pack(pady=5)

        tk.Label(text_fields_frame, text="Данные акта осмотра").pack(anchor=tk.W)
        self.txb_inspection = ttk.Entry(text_fields_frame, width=50)
        self.txb_inspection.pack(pady=5)

        # Текстовое поле внизу
        large_text_field = tk.Text(right_frame, width=80, height=10)
        large_text_field.pack(pady=10)

        # Фрейм для фото
        self.photo_window = tk.Toplevel(self.root)
        self.photo_window.withdraw()
        self.photo_label = tk.Label(self.photo_window)
        self.photo_label.pack()

        ttk.Button(self.photo_window, text="Предыдущее", command=self.prev_photo).pack(side=tk.LEFT, padx=10, pady=10)
        ttk.Button(self.photo_window, text="Следующее", command=self.next_photo).pack(side=tk.RIGHT, padx=10, pady=10)

        # Загрузка списка домов
        self.load_addresses()

    def load_addresses(self):
        try:
            df = pd.read_excel("data/Fund.xlsx", sheet_name="Адресный перечень")
            for _, row in df.iterrows():
                address = str(row.iloc[0])
                self.listbox.insert(tk.END, address)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить список домов:\n{e}")

    def get_selected_address(self):
        selected = self.listbox.curselection()
        if selected:
            self.selected_address = self.listbox.get(selected[0])
            return True
        else:
            messagebox.showwarning("Внимание", "Выберите адрес дома из списка.")
            return False

    def open_passport(self):
        if not self.get_selected_address():
            return

        filename = f"Паспорт готовности к эксплуатации ({self.selected_address}).xlsx"
        self.passport_path = os.path.join("templates", filename)

        if os.path.exists(self.passport_path):
            try:
                self.excel_app = win32.Dispatch("Excel.Application")
                self.excel_app.Visible = False
                self.passport_wb = self.excel_app.Workbooks.Open(os.path.abspath(self.passport_path))
                self.ws = self.passport_wb.Sheets(1)
                messagebox.showinfo("Открыто", f"Паспорт открыт: {filename}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть паспорт:\n{e}")
        else:
            messagebox.showerror("Ошибка", f"Файл не найден:\n{self.passport_path}")

    def fill_data_from_inspection(self):
        if not hasattr(self, 'ws'):
            messagebox.showwarning("Внимание", "Сначала откройте паспорт.")
            return

        try:
            today = datetime.now().strftime("%d %B %Y")
            self.ws.Range("BA7").Value = today  # Дата
            self.ws.Range("V24").Value = "АО \"Мосэнергосбыт\""
            self.ws.Range("AG26").Value = "ООО \"Пример организация\""
            self.ws.Range("V27").Value = "АО \"Мосэнергосбыт\""

            inspection_file = f"Акт общего осмотра ({self.selected_address}).xlsx"
            inspection_path = os.path.join("data", inspection_file)

            if os.path.exists(inspection_path):
                ins_excel = win32.Dispatch("Excel.Application")
                ins_excel.Visible = False
                ins_wb = ins_excel.Workbooks.Open(os.path.abspath(inspection_path))
                ins_ws = ins_wb.Sheets(1)

                value = ins_ws.Range("AK130").Value
                self.ws.Range("AS70").Value = value

                ins_wb.Close(SaveChanges=False)
                ins_excel.Quit()

                messagebox.showinfo("Данные", "Поля успешно заполнены.")
            else:
                messagebox.showwarning("Внимание", f"Акт осмотра не найден: {inspection_file}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка заполнения данных:\n{e}")

    def save_and_close_passport(self):
        if not self.passport_wb or not self.excel_app:
            messagebox.showwarning("Внимание", "Нет открытого паспорта для сохранения.")
            return

        try:
            self.passport_wb.Save()
            self.passport_wb.Close(SaveChanges=False)
            self.excel_app.Quit()
            del self.passport_wb, self.excel_app
            self.passport_wb = self.excel_app = None
            messagebox.showinfo("Сохранено", "Файл успешно сохранён и закрыт.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    def show_photos(self):
        if not self.get_selected_address():
            return

        photo_dir = os.path.join("photos", self.selected_address)
        if not os.path.exists(photo_dir):
            messagebox.showwarning("Внимание", f"Папка с фото не найдена:\n{photo_dir}")
            return

        self.photos = [os.path.join(photo_dir, f) for f in os.listdir(photo_dir) if f.lower().endswith((".png", ".jpg"))]
        if not self.photos:
            messagebox.showwarning("Внимание", "Нет доступных фото для этого дома.")
            return

        self.current_photo_index = 0
        self.display_current_photo()
        self.photo_window.deiconify()

    def display_current_photo(self):
        try:
            img = Image.open(self.photos[self.current_photo_index])
            img.thumbnail((600, 600))
            photo = ImageTk.PhotoImage(img)
            self.photo_label.config(image=photo)
            self.photo_label.image = photo
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить фото:\n{e}")

    def prev_photo(self):
        if self.photos:
            self.current_photo_index = (self.current_photo_index - 1) % len(self.photos)
            self.display_current_photo()

    def next_photo(self):
        if self.photos:
            self.current_photo_index = (self.current_photo_index + 1) % len(self.photos)
            self.display_current_photo()


if __name__ == "__main__":
    root = tk.Tk()
    app = AISToolApp(root)
    root.mainloop()