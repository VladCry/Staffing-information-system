from tkinter import filedialog
import os
import tkinter as tk
from tkinter import messagebox, ttk
import csv
import re

a1, a2, a3, a4 = [], [], [], []
current_data = []
current_filename = ""
validate_code = lambda v: re.fullmatch(r'\d{3}', v) is not None
validate_text = lambda v: re.fullmatch(r'^[a-zA-Zа-яА-ЯёЁ\s-]+$', v) is not None
validate_phone = lambda v: re.fullmatch(r'^\d+$', v) is not None
validate_date = lambda v: re.fullmatch(r'\d{4}-\d{2}-\d{2}', v) is not None
validate_number = lambda v: v.isdigit()


def show_error_and_retry(msg, window, cb):
    messagebox.showerror("Ошибка ввода", msg)
    window.destroy()
    cb()


def import_files():
    global a1, a2, a3, a4
    try:
        for file, var in [('file_post.csv', 'a1'), ('file_schedule.csv', 'a2'), ('file_staff.csv', 'a3'),
                          ('file_department.csv', 'a4')]:
            with open(file, 'r', encoding='utf-8') as f:
                globals()[var] = list(csv.reader(f, delimiter=';'))
        messagebox.showinfo("Импорт", "Файлы успешно считаны!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось считать файлы: {e}")


def view_file(filename):
    global current_data, current_filename
    try:
        if filename == 'file_post.csv':
            content = a1
        elif filename == 'file_schedule.csv':
            content = a2
        elif filename == 'file_staff.csv':
            content = a3
        elif filename == 'file_department.csv':
            content = a4
        else:
            content = []
        current_data = content
        current_filename = filename
        view_window = tk.Toplevel(root)
        view_window.title(filename)
        view_window.geometry("1000x600")
        tree = ttk.Treeview(view_window)
        if content:
            columns = content[0]
            tree["columns"] = columns
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=120, anchor="w")
            for row in content[1:]:
                tree.insert("", "end", values=row)
        scrollbar = ttk.Scrollbar(view_window, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        btn_frame = tk.Frame(view_window)
        btn_frame.pack(side="bottom", fill="x", pady=5)
        tk.Button(btn_frame, text="Добавить запись", command=lambda: add_record(tree, filename)).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Редактировать выделенное", command=lambda: edit_selected(tree, filename)).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Удалить выделенное", command=lambda: delete_selected(tree, filename)).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Сохранить изменения", command=lambda: save_changes(filename)).pack(side="right", padx=5)
        tk.Button(btn_frame, text="Закрыть", command=view_window.destroy).pack(side="right", padx=5)
        scrollbar.pack(side="right", fill="y")
        tree.pack(expand=True, fill="both")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при отображении файла: {e}")


def add_record(tree, filename):
    global current_data
    form_window = tk.Toplevel(root)
    form_window.title(f"Добавление записи в {filename}")
    fields = current_data[0]
    entries = []
    for i, field in enumerate(fields):
        tk.Label(form_window, text=field).grid(row=i, column=0, padx=5, pady=5, sticky="e")
        entry = tk.Entry(form_window, width=40)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entries.append(entry)

    def validate_and_save():
        errors = []
        new_record = []
        for i, (field, entry) in enumerate(zip(fields, entries)):
            value = entry.get()
            new_record.append(value)
            if filename == 'file_post.csv':
                if i == 0:
                    if not validate_code(value): errors.append("Код должности должен быть в формате 000 (3 цифры)")
                else:
                    if not validate_text(value): errors.append("Название должности должно содержать только буквы и пробелы")
            elif filename == 'file_staff.csv':
                if i == 0:
                    if not validate_code(value): errors.append("ID сотрудника должен быть в формате 000 (3 цифры)")
                elif i == 4:
                    if not validate_code(value): errors.append("Код должности должен быть в формате 000 (3 цифры)")
                elif i == 8:
                    if not validate_code(value): errors.append("ID штатного расписания должен быть в формате 000 (3 цифры)")
                elif i == 2:
                    if not validate_date(value): errors.append("Дата рождения должна быть в формате ГГГГ-ММ-ДД")
                elif i == 6:
                    if not validate_phone(value): errors.append("Номер телефона должен содержать только цифры (без +)")
                elif i in [1, 3, 5, 7]:
                    if not value: errors.append(f"Поле {fields[i]} не может быть пустым")
            elif filename == 'file_department.csv':
                if i == 0:
                    if not validate_code(value): errors.append("Код отдела должен быть в формате 000 (3 цифры)")
                elif i in [1, 2]:
                    if not validate_text(value): errors.append(f"Поле {fields[i]} должно содержать только буквы и пробелы")
                elif i == 3:
                    if not validate_phone(value): errors.append("Номер телефона должен содержать только цифры (без +)")
            elif filename == 'file_schedule.csv':
                if i in [0, 1, 2]:
                    if not validate_code(value): errors.append(f"Поле {fields[i]} должно быть в формате 000 (3 цифры)")
                elif i in [3, 4, 5, 6]:
                    if not validate_number(value): errors.append(f"Поле {fields[i]} должно содержать только цифры")
        if errors:
            show_error_and_retry("\n".join(errors), form_window, lambda: add_record(tree, filename))
            return
        current_data.append(new_record)
        update_treeview(tree)
        form_window.destroy()
        messagebox.showinfo("Успех", "Запись добавлена (не забудьте сохранить изменения!)")

    tk.Button(form_window, text="Сохранить", command=validate_and_save).grid(row=len(fields), column=0, pady=10)
    tk.Button(form_window, text="Отмена", command=form_window.destroy).grid(row=len(fields), column=1, pady=10)


def edit_selected(tree, filename):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Предупреждение", "Выберите запись для редактирования")
        return
    item_values = tree.item(selected_item)['values']
    fields = current_data[0]
    form_window = tk.Toplevel(root)
    form_window.title(f"Редактирование записи в {filename}")
    entries = []
    for i, (field, value) in enumerate(zip(fields, item_values)):
        tk.Label(form_window, text=field).grid(row=i, column=0, padx=5, pady=5, sticky="e")
        entry = tk.Entry(form_window, width=40)
        entry.insert(0, value)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entries.append(entry)

    def validate_and_save():
        errors = []
        edited_record = []
        for i, (field, entry) in enumerate(zip(fields, entries)):
            value = entry.get()
            edited_record.append(value)
            if filename == 'file_post.csv':
                if i == 0:
                    if not validate_code(value): errors.append("Код должности должен быть в формате 000 (3 цифры)")
                else:
                    if not validate_text(value): errors.append("Название должности должно содержать только буквы и пробелы")
            elif filename == 'file_staff.csv':
                if i == 0:
                    if not validate_code(value): errors.append("ID сотрудника должен быть в формате 000 (3 цифры)")
                elif i == 1:
                    if not validate_code(value): errors.append("Код должности должен быть в формате 000 (3 цифры)")
                elif i == 2:
                    if not validate_code(value): errors.append("ID штатного расписания должен быть в формате 000 (3 цифры)")
                elif i == 3:
                    if not validate_date(value): errors.append("Дата рождения должна быть в формате ГГГГ-ММ-ДД")
                elif i == 4:
                    if not validate_phone(value): errors.append("Номер телефона должен содержать только цифры (без +)")
                elif i in [5, 6, 7, 8]:
                    if not value: errors.append(f"Поле {fields[i]} не может быть пустым")
            elif filename == 'file_department.csv':
                if i == 0:
                    if not validate_code(value): errors.append("Код отдела должен быть в формате 000 (3 цифры)")
                elif i in [1, 2]:
                    if not validate_text(value): errors.append(f"Поле {fields[i]} должно содержать только буквы и пробелы")
                elif i == 3:
                    if not validate_phone(value): errors.append("Номер телефона должен содержать только цифры (без +)")
            elif filename == 'file_schedule.csv':
                if i in [0, 1, 2]:
                    if not validate_code(value): errors.append(f"Поле {fields[i]} должно быть в формате 000 (3 цифры)")
                elif i in [3, 4, 5, 6]:
                    if not validate_number(value): errors.append(f"Поле {fields[i]} должно содержать только цифры")
        if errors:
            show_error_and_retry("\n".join(errors), form_window, lambda: edit_selected(tree, filename))
            return
        index = tree.index(selected_item) + 1
        current_data[index] = edited_record
        update_treeview(tree)
        form_window.destroy()
        messagebox.showinfo("Успех", "Запись изменена (не забудьте сохранить изменения!)")

    tk.Button(form_window, text="Сохранить", command=validate_and_save).grid(row=len(fields), column=0, pady=10)
    tk.Button(form_window, text="Отмена", command=form_window.destroy).grid(row=len(fields), column=1, pady=10)


def delete_selected(tree, filename):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Предупреждение", "Выберите запись для удаления")
        return
    if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить эту запись?"):
        index = tree.index(selected_item) + 1
        del current_data[index]
        update_treeview(tree)
        messagebox.showinfo("Успех", "Запись удалена (не забудьте сохранить изменения!)")


def update_treeview(tree):
    tree.delete(*tree.get_children())
    for row in current_data[1:]:
        tree.insert("", "end", values=row)


def save_changes(filename):
    global a1, a2, a3, a4, current_data
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerows(current_data)
        if filename == 'file_post.csv':
            a1 = current_data.copy()
        elif filename == 'file_schedule.csv':
            a2 = current_data.copy()
        elif filename == 'file_staff.csv':
            a3 = current_data.copy()
        elif filename == 'file_department.csv':
            a4 = current_data.copy()
        messagebox.showinfo("Успех", "Изменения успешно сохранены!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить изменения: {e}")


def correct():
    edit_window = tk.Toplevel(root)
    edit_window.title("Выбор формы редактирования")
    edit_window.geometry("300x200")
    tk.Label(edit_window, text="Выберите файл для редактирования:", font=('Arial', 12)).pack(pady=10)
    btn_frame = tk.Frame(edit_window)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="Должности", command=lambda: [edit_window.destroy(), view_file('file_post.csv')]).grid(row=0, column=0, padx=5, pady=5)
    tk.Button(btn_frame, text="Сотрудники", command=lambda: [edit_window.destroy(), view_file('file_staff.csv')]).grid(row=1, column=0, padx=5, pady=5)
    tk.Button(btn_frame, text="Отделы", command=lambda: [edit_window.destroy(), view_file('file_department.csv')]).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(btn_frame, text="Расписание", command=lambda: [edit_window.destroy(), view_file('file_schedule.csv')]).grid(row=1, column=1, padx=5, pady=5)


def report():
    if not hasattr(calculate, 'last_result'):
        messagebox.showwarning("Предупреждение", "Сначала выполните расчет потребности в сотрудниках.")
        return
    position = calculate.last_position
    result = calculate.last_result
    header = """
---
2025

Штатное расписание:

"""
    if isinstance(result, int):
        if result > 0:
            report_text = header + f"Должность: {position}\nКоличество вакантных мест: {result}"
        elif result == 0:
            report_text = header + f"Должность: {position}\nНет вакантных мест"
        else:
            report_text = header + f"Должность: {position}\nОшибка: избыток сотрудников ({abs(result)} чел.)"
    else:
        report_text = header + f"Должность: {position}\n{result}"
    filepath = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
        title="Сохранить отчет о вакансиях",
        initialfile=f"Штатное_расписание_{position.replace(' ', '_')}.txt"
    )
    if filepath:
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(report_text)
            messagebox.showinfo("Успех", f"Отчет успешно сохранен:\n{filepath}")
            if messagebox.askyesno("Открыть папку", "Открыть папку с сохраненным отчетом?"):
                folder_path = os.path.dirname(filepath)
                os.startfile(folder_path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчет:\n{str(e)}")


def calculate():
    def perform_calculation():
        position = position_var.get()
        if position:
            result = chet(position)
            result_label.config(text=f"Результат: {result}")
            calculate.last_position = position
            calculate.last_result = result
            report_button = tk.Button(calc_window, text="Сохранить отчет", command=report, bg="#4CAF50", fg="white")
            report_button.pack(pady=10)
        else:
            messagebox.showwarning("Предупреждение", "Выберите должность.")

    calc_window = tk.Toplevel(root)
    calc_window.title("Расчет потребности в сотрудниках")
    calc_window.geometry("400x300")
    tk.Label(calc_window, text="Выберите должность:", font=('Arial', 12)).pack(pady=10)

    positions = []
    try:
        with open('file_post.csv', 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')
            next(reader)
            positions = [row[1] for row in reader if len(row) > 1]
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить должности: {e}")
        calc_window.destroy()
        return

    if not positions:
        messagebox.showwarning("Предупреждение", "Список должностей пуст!")
        calc_window.destroy()
        return

    position_var = tk.StringVar(calc_window)
    position_var.set(positions[0])
    position_menu = tk.OptionMenu(calc_window, position_var, *positions)
    position_menu.pack(pady=10)

    calculate_button = tk.Button(calc_window, text="Посчитать", command=perform_calculation)
    calculate_button.pack(pady=20)

    result_label = tk.Label(calc_window, text="Результат: ")
    result_label.pack(pady=10)


def chet(post):
    id_pos = None
    id_schedule = None
    kolvo = 0
    count = 0

    for row in a1[1:]:
        if row[1] == post:
            id_pos = row[0]
            break

    if id_pos is not None:
        for row in a2[1:]:
            if row[0] == id_pos:
                id_schedule = row[1]
                kolvo = row[3]
                break

        if id_schedule is not None:
            for row in a3[1:]:
                if row[-1] == id_schedule:
                    count += 1

            itog = int(kolvo) - count
            return itog

    return "Данные не найдены"


def main():
    global root
    root = tk.Tk()
    root.title("Информационная система штатного расписания")
    root.geometry("800x600")

    menu = tk.Menu(root)
    root.config(menu=menu)

    import_menu = tk.Menu(menu)
    menu.add_cascade(label="Импорт", menu=import_menu)
    import_menu.add_command(label="Считать файлы", command=import_files)

    menu.add_command(label="Редактирование", command=correct)

    view_menu = tk.Menu(menu)
    menu.add_cascade(label="Просмотр", menu=view_menu)
    view_menu.add_command(label="Должности", command=lambda: view_file('file_post.csv'))
    view_menu.add_command(label="Расписание", command=lambda: view_file('file_schedule.csv'))
    view_menu.add_command(label="Сотрудники", command=lambda: view_file('file_staff.csv'))
    view_menu.add_command(label="Отдел", command=lambda: view_file('file_department.csv'))

    menu.add_command(label="Расчет", command=calculate)

    root.mainloop()


if __name__ == "__main__":
    main()