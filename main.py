import tkinter as tk  # стандартна бібліотека для створення GUI в Python
from datetime import datetime  # використовується для роботи з датами
from tkinter import font as tkfont, \
    simpledialog  # використовуються для більш детального управління елементами інтерфейсу
import pandas as pd  # використовується для ефективного управління та маніпуляції даними


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Excel App")
        self.grid_data = pd.DataFrame(index=range(1, 6), columns=list('ABCDE')).fillna('')
        self.selected_cell = None
        self.create_grid()

    # Робимо грід 5 на 5
    def create_grid(self):
        bold_font = tkfont.Font(weight="bold")

        for c in range(5):
            tk.Label(self.root, text=chr(65 + c), font=bold_font).grid(row=0, column=c + 1)
            tk.Label(self.root, text=str(c + 1), font=bold_font).grid(row=c + 1, column=0)

        self.cells = {}
        normal_font = tkfont.Font(weight="normal")

        for r in range(1, 6):
            for c in range(1, 6):
                cell = tk.Entry(self.root, justify='right', font=normal_font)
                cell.grid(row=r, column=c, sticky='nsew')
                cell.insert(0, self.grid_data.iloc[r - 1, c - 1])
                cell.bind('<FocusIn>', lambda e, r=r, c=c: self.select_cell(r, c, e.widget)) # Ці помилки потрібні.
                self.cells[(r, c)] = cell

        # Робимо кнопки
        self.root.grid_columnconfigure(tuple(range(6)), weight=1)
        self.root.grid_rowconfigure(tuple(range(6)), weight=1)

        self.color_button = tk.Button(self.root, text='Змінити колір', command=self.toggle_color)
        self.color_button.grid(row=6, column=0, columnspan=2, sticky='nsew')

        self.bold_button = tk.Button(self.root, text='Жирний текст', command=self.toggle_bold)
        self.bold_button.grid(row=6, column=2, columnspan=2, sticky='nsew')

        self.format_button = tk.Button(self.root, text='Форматувати як валюту', command=self.toggle_format_currency)
        self.format_button.grid(row=6, column=4, columnspan=2, sticky='nsew')

        self.format_date_button = tk.Button(self.root, text='Форматувати дату', command=self.format_date)
        self.format_date_button.grid(row=6, column=6, columnspan=2, sticky='nsew')

    # Функція вибору клітинки
    def select_cell(self, r, c, widget):
        self.selected_cell = (r, c, widget)

    # Функція виділення коліром клітинки
    def toggle_color(self):
        cell = self.selected_cell[2]
        current_color = cell.cget('bg')
        new_color = 'yellow' if current_color != 'yellow' else 'white'
        cell.config(bg=new_color)

    # Функція яка робить текст жирним (можуть засудити за шейпшеймінг)
    def toggle_bold(self):
        cell = self.selected_cell[2]
        current_font = tkfont.Font(font=cell['font'])
        is_bold = current_font.actual()['weight'] == 'bold'
        new_font = tkfont.Font(weight="bold" if not is_bold else "normal")
        cell.config(font=new_font)

    # Функція яка переводить числове значення у долар
    def toggle_format_currency(self):
        cell = self.selected_cell[2]
        value = cell.get()
        if '$' in value:
            new_value = value.replace('$', '').replace(',', '').strip()
        else:
            try:
                value = float(value)
                new_value = '${:,.2f}'.format(value)
            except ValueError:
                return  # не число, ігноруємо
        cell.delete(0, tk.END)
        cell.insert(0, new_value)

    # Функція яка перетворює формат дати у нормальний вигляд
    def format_date(self):
        cell = self.selected_cell[2]
        value = cell.get()
        try:
            date_obj = datetime.strptime(value, '%Y-%m-%d')
            format_choice = simpledialog.askstring("Формат дати",
                                                   "Виберіть формат:\n1 для '01 November 2023'\n2 для '01.12.2023'")
            if format_choice == '1':
                new_value = date_obj.strftime('%d %B %Y')
            elif format_choice == '2':
                new_value = date_obj.strftime('%d.%m.%Y')
            else:
                return  # невідомий формат, ігноруємо
            cell.delete(0, tk.END)
            cell.insert(0, new_value)
        except ValueError:
            return  # не вдається розпізнати дату, ігноруємо


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
