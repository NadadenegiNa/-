import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
from pathlib import Path
from datetime import datetime
from PIL import Image, ImageTk

class RealEstateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Счастье РИЭЛТОРА")
        self.root.geometry("1300x850")
        
        # Определяем путь к папке документов пользователя
        documents_path = Path.home() / "Documents"
        self.filename = documents_path / "real_estate_database.csv"
        
        # Словарь для хранения состояний чекбоксов
        self.checkbox_vars = {}
        
        # Загрузка данных
        self.load_data()
        
        # Создание интерфейса
        self.create_widgets()
        
    def load_data(self):
        if os.path.exists(self.filename):
            self.df = pd.read_csv(self.filename, encoding='utf-8')
            # Преобразуем дату в формат datetime
            if 'Дата занесения в базу' in self.df.columns:
                self.df['Дата занесения в базу'] = pd.to_datetime(self.df['Дата занесения в базу'])
        else:
            self.df = pd.DataFrame(columns=[
                "Дата занесения в базу", "Объект", "Адрес", "Цена", "Площадь", "Комнаты", "Жилая", 
                "Кухня", "Санузел", "Этаж/этажность", "Участок", "Дом", 
                "Высота потолков", "Год постройки", "Сделка", "Основание владения",
                "Срок владения", "Количество собственников", "Контакт ФИО", 
                "Телефон", "Примечание"
            ])
    
    def save_data(self):
        # Создаем папку, если она не существует
        os.makedirs(os.path.dirname(self.filename), exist_ok=True)
        self.df.to_csv(self.filename, index=False, encoding='utf-8')
    
    def create_widgets(self):
        # Декоративный заголовок только с домиками
        decoration_frame = ttk.Frame(self.root)
        decoration_frame.pack(pady=5)
        
        # Заголовок только с символами недвижимости
        ttk.Label(decoration_frame, text="🏠 🏢 🏡 🏘️ 🏘️ 🏡 🏢 🏠", 
                 font=("Arial", 16, "bold"), foreground="darkblue").pack()
        
        # Основной фрейм для полей ввода и картинки
        main_input_frame = ttk.Frame(self.root)
        main_input_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Фрейм для полей ввода
        input_frame = ttk.Frame(main_input_frame)
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Создаем Canvas и Scrollbar для полей ввода
        canvas = tk.Canvas(input_frame)
        scrollbar = ttk.Scrollbar(input_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Поля ввода в правильном порядке
        fields = [
            ("Дата занесения в базу", "date"),
            ("Объект", "combobox", ["комната", "1-ккв", "2-ккв", "3-ккв", "4-ккв", "дача", "дом", "участок"]),
            ("Адрес", "entry"),
            ("Цена", "entry"),
            ("Площадь", "entry"),
            ("Комнаты", "entry"),
            ("Жилая", "entry"),
            ("Кухня", "entry"),
            ("Санузел", "combobox", ["раздельный", "совмещенный"]),
            ("Этаж/этажность", "entry"),
            ("Участок", "entry"),
            ("Дом", "combobox", ["нет", "монолит", "панельный", "кирпичный", "бревно", "каркасный"]),
            ("Высота потолков", "entry"),
            ("Год постройки", "entry"),
            ("Сделка", "combobox", ["прямая продажа", "альтернатива"]),
            ("Основание владения", "combobox", ["ДКП", "ДДУ", "Наследство", "Дарение", "Приватизация"]),
            ("Срок владения", "entry"),
            ("Количество собственников", "entry"),
            ("Контакт ФИО", "entry"),
            ("Телефон", "entry"),
            ("Примечание", "text")
        ]
        
        self.entries = {}
        for i, (field, field_type, *options) in enumerate(fields):
            # Создаем фрейм для каждой строки
            row_frame = ttk.Frame(scrollable_frame)
            row_frame.grid(row=i, column=0, sticky='ew', padx=5, pady=2)
            
            label = ttk.Label(row_frame, text=field, width=20)
            label.pack(side=tk.LEFT, padx=5)
            
            if field_type == "combobox":
                entry = ttk.Combobox(row_frame, values=options[0], state="readonly", width=40)
            elif field_type == "text":
                entry = tk.Text(row_frame, height=3, width=40)
            elif field_type == "date":
                entry = ttk.Entry(row_frame, width=40)
                # Устанавливаем текущую дату по умолчанию
                entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
            else:
                entry = ttk.Entry(row_frame, width=40)
            
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            self.entries[field] = entry
        
        # Упаковка canvas и scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Фрейм для картинки
        image_frame = ttk.Frame(main_input_frame)
        image_frame.pack(side=tk.RIGHT, padx=10)
        
        # Загрузка и отображение картинки
        try:
            image_path = "C:\\Users\\Pro\\Desktop\\#abracrocodaber\\Титульная картинка.jpg"
            if os.path.exists(image_path):
                image = Image.open(image_path)
                # Масштабируем изображение до разумного размера
                image = image.resize((300, 400), Image.Resampling.LANCZOS)
                self.photo = ImageTk.PhotoImage(image)
                image_label = ttk.Label(image_frame, image=self.photo)
                image_label.pack()
            else:
                # Если изображение не найдено, показываем заглушку
                placeholder = ttk.Label(image_frame, text="Изображение не найдено", width=30, height=15)
                placeholder.pack()
        except Exception as e:
            print(f"Ошибка загрузки изображения: {e}")
            placeholder = ttk.Label(image_frame, text="Ошибка загрузки изображения", width=30, height=15)
            placeholder.pack()
        
        # Фрейм для кнопок с декоративными элементами
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)
        
        # Кнопка "Добавить объект" с домиком
        ttk.Label(button_frame, text="🏠", font=("Arial", 14)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Добавить объект", command=self.add_property).pack(side=tk.LEFT, padx=5)
        
        # Кнопка "Удалить объект" с домиком
        ttk.Label(button_frame, text="🗑️", font=("Arial", 14)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Удалить объект", command=self.delete_property).pack(side=tk.LEFT, padx=5)
        
        # Кнопка "Распечатать" с домиком
        ttk.Label(button_frame, text="📊", font=("Arial", 14)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Распечатать", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        
        # Фрейм для таблицы с прокруткой (уменьшенный размер)
        table_frame = ttk.Frame(self.root, relief="solid", borderwidth=1)
        table_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5, ipady=5)
        
        # Заголовок таблицы
        table_label = ttk.Label(table_frame, text="База объектов недвижимости", font=("Arial", 12, "bold"))
        table_label.pack(pady=5)
        
        # Создаем таблицу с горизонтальной и вертикальной прокруткой
        # Используем правильный порядок столбцов
        columns_order = ["Выбор"] + [field[0] for field in fields]
        self.tree = ttk.Treeview(table_frame, columns=columns_order, show="headings", height=8)  # Уменьшили высоту
        
        # Настройка стиля для таблицы с границами
        style = ttk.Style()
        style.configure("Treeview", 
                       background="white",
                       foreground="black",
                       rowheight=25,
                       fieldbackground="white",
                       borderwidth=1,
                       relief="solid")
        style.configure("Treeview.Heading", 
                       background="lightblue",
                       foreground="black",
                       relief="raised",
                       borderwidth=1)
        
        # Настройка столбцов в правильном порядке
        self.tree.heading("Выбор", text="Выбор")
        self.tree.column("Выбор", width=50, minwidth=50)
        
        # Настраиваем ширину столбцов для лучшего отображения
        column_widths = {
            "Дата занесения в базу": 100,
            "Объект": 70,
            "Адрес": 120,
            "Цена": 80,
            "Площадь": 60,
            "Комнаты": 70,
            "Жилая": 60,
            "Кухня": 60,
            "Санузел": 90,
            "Этаж/этажность": 90,
            "Участок": 70,
            "Дом": 90,
            "Высота потолков": 90,
            "Год постройки": 90,
            "Сделка": 90,
            "Основание владения": 110,
            "Срок владения": 90,
            "Количество собственников": 130,
            "Контакт ФИО": 120,
            "Телефон": 100,
            "Примечание": 150
        }
        
        for col in columns_order[1:]:  # Пропускаем столбец "Выбор"
            self.tree.heading(col, text=col)
            width = column_widths.get(col, 100)
            self.tree.column(col, width=width, minwidth=50)
        
        # Увеличиваем ширину для столбцов с длинным текстом
        self.tree.column("Контакт ФИО", width=150)
        self.tree.column("Телефон", width=120)
        self.tree.column("Примечание", width=200)
        
        # Вертикальная прокрутка
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Горизонтальная прокрутка
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Размещение элементов
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Привязываем обработчик клика для чекбоксов
        self.tree.bind('<Button-1>', self.on_tree_click)
        
        # Привязываем двойной клик для редактирования
        self.tree.bind('<Double-1>', self.on_double_click)
        
        self.update_table()
    
    def on_tree_click(self, event):
        # Определяем, по какому столбцу и элементу был клик
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
            
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        # Если клик был по столбцу "Выбор"
        if column == "#1" and item:
            # Получаем текущее состояние
            current_values = self.tree.item(item, 'values')
            if not current_values:
                return
                
            # Создаем список значений
            values_list = list(current_values)
            
            # Изменяем состояние чекбокса
            if values_list[0] == "☐":
                values_list[0] = "☑"
                # Выделяем строку
                self.tree.selection_add(item)
            else:
                values_list[0] = "☐"
                # Снимаем выделение
                self.tree.selection_remove(item)
            
            # Обновляем значения в таблице
            self.tree.item(item, values=tuple(values_list))
    
    def on_double_click(self, event):
        # Получаем выбранный элемент
        item = self.tree.selection()
        if not item:
            return
            
        item = item[0]
        values = self.tree.item(item, 'values')
        
        # Пропускаем столбец с чекбоксом
        if len(values) > 1:
            self.edit_property(item, values[1:])
    
    def edit_property(self, item, values):
        # Создаем окно редактирования
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактирование объекта")
        edit_window.geometry("600x700")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        # Получаем индекс объекта в DataFrame
        index = self.tree.index(item)
        
        # Создаем поля для редактирования
        fields = [
            ("Дата занесения в базу", "date"),
            ("Объект", "combobox", ["комната", "1-ккв", "2-ккв", "3-ккв", "4-ккв", "дача", "дом", "участок"]),
            ("Адрес", "entry"),
            ("Цена", "entry"),
            ("Площадь", "entry"),
            ("Комнаты", "entry"),
            ("Жилая", "entry"),
            ("Кухня", "entry"),
            ("Санузел", "combobox", ["раздельный", "совмещенный"]),
            ("Этаж/этажность", "entry"),
            ("Участок", "entry"),
            ("Дом", "combobox", ["нет", "монолит", "панельный", "кирпичный", "бревно", "каркасный"]),
            ("Высота потолков", "entry"),
            ("Год постройки", "entry"),
            ("Сделка", "combobox", ["прямая продажа", "альтернатива"]),
            ("Основание владения", "combobox", ["ДКП", "ДДУ", "Наследство", "Дарение", "Приватизация"]),
            ("Срок владения", "entry"),
            ("Количество собственников", "entry"),
            ("Контакт ФИО", "entry"),
            ("Телефон", "entry"),
            ("Примечание", "text")
        ]
        
        edit_entries = {}
        for i, (field, field_type, *options) in enumerate(fields):
            # Создаем фрейм для каждой строки
            row_frame = ttk.Frame(edit_window)
            row_frame.pack(fill=tk.X, padx=10, pady=2)
            
            label = ttk.Label(row_frame, text=field, width=20)
            label.pack(side=tk.LEFT, padx=5)
            
            if field_type == "combobox":
                entry = ttk.Combobox(row_frame, values=options[0], state="readonly", width=40)
                if i < len(values):
                    entry.set(values[i])
            elif field_type == "text":
                entry = tk.Text(row_frame, height=3, width=40)
                if i < len(values):
                    entry.insert("1.0", values[i])
            elif field_type == "date":
                entry = ttk.Entry(row_frame, width=40)
                if i < len(values):
                    entry.insert(0, values[i])
            else:
                entry = ttk.Entry(row_frame, width=40)
                if i < len(values):
                    entry.insert(0, values[i])
            
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            edit_entries[field] = entry
        
        # Фрейм для кнопок
        button_frame = ttk.Frame(edit_window)
        button_frame.pack(pady=10)
        
        # Кнопка "ОК"
        ttk.Button(button_frame, text="ОК", command=lambda: self.save_edit(index, edit_entries, edit_window)).pack(side=tk.LEFT, padx=5)
        
        # Кнопка "Отмена"
        ttk.Button(button_frame, text="Отмена", command=edit_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def save_edit(self, index, edit_entries, edit_window):
        try:
            # Собираем данные из полей редактирования
            updated_row = {}
            for field, entry in edit_entries.items():
                if isinstance(entry, tk.Text):
                    value = entry.get("1.0", tk.END).strip()
                else:
                    value = entry.get()
                
                # Обработка специальных полей
                if field == "Участок" and value and "соток" not in value:
                    value = f"{value} соток"
                
                updated_row[field] = value
            
            # Обновляем строку в DataFrame
            for field, value in updated_row.items():
                self.df.at[index, field] = value
            
            # Сохраняем изменения
            self.save_data()
            self.update_table()
            
            # Закрываем окно редактирования
            edit_window.destroy()
            
            messagebox.showinfo("Успех", "Объект успешно отредактирован!")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при редактировании: {str(e)}")
    
    def add_property(self):
        try:
            new_row = {}
            for field, entry in self.entries.items():
                if isinstance(entry, tk.Text):
                    value = entry.get("1.0", tk.END).strip()
                else:
                    value = entry.get()
                
                # Обработка специальных полей
                if field == "Участок" and value:
                    value = f"{value} соток"
                elif field == "Дата занесения в базу" and not value:
                    # Устанавливаем текущую дату, если поле пустое
                    value = datetime.now().strftime("%d.%m.%Y")
                
                new_row[field] = value
            
            # Добавляем новую строку в DataFrame
            self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
            self.save_data()
            self.update_table()
            self.clear_form()
            messagebox.showinfo("Успех", "Объект успешно добавлен!")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении: {str(e)}")
    
    def delete_property(self):
        try:
            # Получаем все элементы таблицы
            items = self.tree.get_children()
            indices_to_delete = []
            
            # Находим индексы отмеченных строк
            for i, item in enumerate(items):
                values = self.tree.item(item, 'values')
                if values and values[0] == "☑":
                    indices_to_delete.append(i)
            
            if not indices_to_delete:
                messagebox.showwarning("Предупреждение", "Не выбрано ни одного объекта для удаления.")
                return
            
            # Подтверждение удаления
            confirm = messagebox.askyesno(
                "Подтверждение удаления", 
                f"Вы уверены, что хотите удалить {len(indices_to_delete)} объектов?"
            )
            
            if not confirm:
                return
            
            # Удаляем строки из DataFrame (в обратном порядке, чтобы индексы не сдвигались)
            for index in sorted(indices_to_delete, reverse=True):
                self.df = self.df.drop(index).reset_index(drop=True)
            
            # Сохраняем изменения и обновляем таблицу
            self.save_data()
            self.update_table()
            
            messagebox.showinfo("Успех", f"Удалено объектов: {len(indices_to_delete)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении: {str(e)}")
    
    def update_table(self):
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Заполняем таблицу данными в правильном порядке
        for _, row in self.df.iterrows():
            # Создаем список значений в правильном порядке
            values = ["☐"]  # Чекбокс
            for field in self.entries.keys():
                if field in row:
                    # Для полей с длинным текстом добавляем переносы
                    value = str(row[field])
                    if field in ["Контакт ФИО", "Телефон", "Примечание"] and len(value) > 20:
                        # Добавляем переносы каждые 20 символов
                        value = '\n'.join([value[i:i+20] for i in range(0, len(value), 20)])
                    values.append(value)
                else:
                    values.append("")
            
            self.tree.insert("", tk.END, values=values)
    
    def clear_form(self):
        for field, entry in self.entries.items():
            if isinstance(entry, tk.Text):
                entry.delete("1.0", tk.END)
            else:
                entry.delete(0, tk.END)
                # Для даты устанавливаем текущую дату по умолчанию
                if field == "Дата занесения в базу":
                    entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
    
    def export_to_excel(self):
        try:
            # Создаем копию данных для экспорта
            export_df = self.df.copy()
            
            # Сортируем по объекту в заданном порядке
            order = ["комната", "1-ккв", "2-ккв", "3-ккв", "4-ккв", "дача", "дом", "участок"]
            export_df['Объект'] = pd.Categorical(export_df['Объект'], categories=order, ordered=True)
            export_df = export_df.sort_values('Объект')
            
            # Переименовываем столбцы согласно требованиям
            column_mapping = {
                "Объект": "Объект",
                "Адрес": "адрес",
                "Цена": "цена",
                "Площадь": "S",
                "Комнаты": "комнаты",
                "Жилая": "жил",
                "Кухня": "Кух",
                "Санузел": "с/у",
                "Этаж/этажность": "этаж",
                "Участок": "участок",
                "Дом": "дом",
                "Год постройки": "год",
                "Контакт ФИО": "контакт",
                "Телефон": "телефон",
                "Примечание": "ВАЖНО!"
            }
            
            # Оставляем только нужные столбцы и переименовываем их
            export_df = export_df.rename(columns=column_mapping)
            columns_to_export = list(column_mapping.values())
            # Убедимся, что все столбцы существуют
            existing_columns = [col for col in columns_to_export if col in export_df.columns]
            export_df = export_df[existing_columns]
            
            # Сохраняем в Excel
            documents_path = Path.home() / "Documents"
            excel_filename = documents_path / "real_estate_export.xlsx"
            export_df.to_excel(excel_filename, index=False)
            
            messagebox.showinfo("Успех", f"Данные экспортированы в {excel_filename}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RealEstateApp(root)
    root.mainloop()