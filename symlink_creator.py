import os
import ctypes
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import threading

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def paste_from_clipboard(entry):
    try:
        clipboard_text = root.clipboard_get()
        if clipboard_text.strip():
            entry.delete(0, tk.END)
            entry.insert(0, clipboard_text.strip())
            if entry == source_entry:
                update_symlink_name()
    except:
        messagebox.showwarning("Ошибка", "Не удалось вставить текст")

def clear_entry(entry):
    entry.delete(0, tk.END)
    if entry == source_entry:
        update_symlink_name()

def browse_folder(entry):
    folder = filedialog.askdirectory()
    if folder:
        entry.delete(0, tk.END)
        entry.insert(0, folder)
        if entry == source_entry:
            update_symlink_name()

def update_symlink_name(*args):
    if not manual_name_var.get():
        source = source_entry.get().strip()
        if source:
            symlink_name = os.path.basename(os.path.normpath(source))
            name_entry.delete(0, tk.END)
            name_entry.insert(0, symlink_name)

def create_symlink():
    source = source_entry.get().strip()
    target_folder = target_entry.get().strip()
    symlink_name = name_entry.get().strip()

    if not source or not target_folder or not symlink_name:
        messagebox.showerror("Ошибка", "Заполните все поля")
        return

    try:
        source = os.path.normpath(source)
        target_folder = os.path.normpath(target_folder)
        target = os.path.join(target_folder, symlink_name)

        if not os.path.exists(source):
            messagebox.showerror("Ошибка", f"Исходная папка/файл не существует:\n{source}")
            return

        if os.path.exists(target):
            messagebox.showerror("Ошибка", f"Целевой путь уже существует:\n{target}")
            return

        if not is_admin():
            messagebox.showwarning("Ошибка", "Требуются права администратора")
            return

        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        if os.path.isdir(source):
            os.symlink(source, target, target_is_directory=True)
        else:
            os.symlink(source, target)

        messagebox.showinfo("Успех", f"Симлинк создан:\n{target} → {source}")

    except PermissionError:
        messagebox.showerror("Ошибка", "Недостаточно прав для выполнения операции")
    except OSError as e:
        messagebox.showerror("Ошибка", f"Не удалось создать симлинк:\n{str(e)}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Неизвестная ошибка:\n{str(e)}")

def get_available_drives():
    drives = []
    try:
        for letter in range(65, 91):
            drive = chr(letter) + ":\\"
            if os.path.exists(drive):
                drives.append(drive)
    except:
        messagebox.showwarning("Ошибка", "Не удалось получить список дисков")
    return drives

def search_files():
    drive = drive_var.get()
    if not drive:
        messagebox.showerror("Ошибка", "Выберите диск для поиска")
        return

    search_extensions = []
    if dll_var.get():
        search_extensions.append('.dll')
    if vst3_var.get():
        search_extensions.append('.vst3')
    if not search_extensions:
        messagebox.showerror("Ошибка", "Выберите хотя бы один тип файла (*.dll или *.vst3)")
        return

    for item in tree.get_children():
        tree.delete(item)

    found_files.clear()
    check_vars.clear()
    progress_label.config(text="Поиск начат...")
    search_button.config(state="disabled")
    cancel_button.config(state="normal")
    global search_cancelled
    search_cancelled = False

    def search_recursive(path):
        if search_cancelled:
            return
        try:
            for entry in os.scandir(path):
                if search_cancelled:
                    return
                if entry.is_dir() and entry.name in ("Windows", "Program Files", "Program Files (x86)", "System Volume Information", "$RECYCLE.BIN"):
                    continue
                if entry.is_file() and entry.name.lower().endswith(tuple(search_extensions)):
                    found_files.append(entry.path)
                    root.after(0, lambda: progress_label.config(text=f"Найдено файлов: {len(found_files)}"))
                    if len(found_files) >= 1000:
                        return
                elif entry.is_dir():
                    search_recursive(entry.path)
        except (PermissionError, OSError):
            pass

    def search_thread():
        try:
            search_recursive(drive)
            root.after(0, lambda: progress_label.config(text=f"Поиск завершён. Найдено файлов: {len(found_files)}"))
            root.after(0, lambda: populate_treeview())
            if len(found_files) >= 1000:
                root.after(0, lambda: messagebox.showwarning("Предупреждение", "Достигнуто ограничение в 1000 файлов. Поиск остановлен."))
        except Exception as e:
            root.after(0, lambda: progress_label.config(text="Поиск завершён с ошибкой"))
            root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка при поиске:\n{str(e)}"))
        finally:
            root.after(0, lambda: search_button.config(state="normal"))
            root.after(0, lambda: cancel_button.config(state="disabled"))

    def populate_treeview():
        for file_path in found_files:
            parent_folder = os.path.basename(os.path.dirname(file_path))
            var = tk.BooleanVar(value=False)
            check_vars.append(var)
            tree.insert("", "end", values=("[ ]", file_path, parent_folder), tags=(str(len(check_vars)-1),))

    threading.Thread(target=search_thread, daemon=True).start()

def cancel_search():
    global search_cancelled
    search_cancelled = True
    progress_label.config(text="Поиск отменён")
    search_button.config(state="normal")
    cancel_button.config(state="disabled")

def toggle_check(event):
    col = tree.identify_column(event.x)
    if col != "#1":  # Проверяем, что щелчок был по колонке Check
        return
    item = tree.identify_row(event.y)
    if item:
        tag = tree.item(item, "tags")[0]
        var_index = int(tag)
        var = check_vars[var_index]
        var.set(not var.get())
        tree.item(item, values=("[x]" if var.get() else "[ ]", tree.item(item, "values")[1], tree.item(item, "values")[2]))

def select_all():
    for var in check_vars:
        var.set(True)
    for item in tree.get_children():
        tree.item(item, values=("[x]", tree.item(item, "values")[1], tree.item(item, "values")[2]))

def deselect_all():
    for var in check_vars:
        var.set(False)
    for item in tree.get_children():
        tree.item(item, values=("[ ]", tree.item(item, "values")[1], tree.item(item, "values")[2]))

def create_selected_symlinks():
    target_folder = search_target_entry.get().strip()
    if not target_folder:
        messagebox.showerror("Ошибка", "Укажите целевую папку")
        return

    selected_items = []
    for item in tree.get_children():
        tag = tree.item(item, "tags")[0]
        var_index = int(tag)
        if check_vars[var_index].get():
            selected_items.append(tree.item(item))

    if not selected_items:
        messagebox.showerror("Ошибка", "Выберите хотя бы один файл")
        return

    created_count = 0
    try:
        target_folder = os.path.normpath(target_folder)
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        for item in selected_items:
            source = item["values"][1]  # file_path
            parent_folder = item["values"][2]  # parent_folder
            file_name = os.path.splitext(os.path.basename(source))[0]
            symlink_name = f"{parent_folder}_{file_name}"
            target = os.path.join(target_folder, symlink_name)

            counter = 1
            base_symlink_name = symlink_name
            while os.path.exists(target):
                symlink_name = f"{base_symlink_name}_{counter}"
                target = os.path.join(target_folder, symlink_name)
                counter += 1

            if not is_admin():
                messagebox.showwarning("Ошибка", "Требуются права администратора")
                return

            os.symlink(source, target)
            created_count += 1

        if created_count > 0:
            messagebox.showinfo("Успех", f"Успешно создано {created_count} симлинков")
        else:
            messagebox.showwarning("Предупреждение", "Ни один симлинк не был создан")

    except PermissionError:
        messagebox.showerror("Ошибка", f"Недостаточно прав для выполнения операции. Создано {created_count} симлинков")
    except OSError as e:
        messagebox.showerror("Ошибка", f"Не удалось создать симлинк: {str(e)}. Создано {created_count} симлинков")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Неизвестная ошибка: {str(e)}. Создано {created_count} симлинков")

# Создаем главное окно
root = tk.Tk()
root.title("Создание симлинков")
root.geometry("1010x650")

# Стили
font_bold = ('Arial', 10, 'bold')
font_normal = ('Arial', 10)
pad_options = {'padx': 5, 'pady': 5, 'sticky': 'we'}

# Настройка стиля для Treeview
style = ttk.Style()
style.configure("Treeview", font=font_normal, rowheight=25)
style.map("Treeview", background=[('selected', '#a6d8ff')])

# Создаем вкладки
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Вкладка "Добавление"
add_frame = ttk.Frame(notebook)
notebook.add(add_frame, text="Добавление")

# Поле исходной папки
tk.Label(add_frame, text="Исходная папка/файл:", font=font_bold).grid(row=0, column=0, **pad_options)
source_entry = tk.Entry(add_frame, width=80, font=font_normal)
source_entry.grid(row=0, column=1, **pad_options)
tk.Button(add_frame, text="Обзор", command=lambda: browse_folder(source_entry)).grid(row=0, column=2, **pad_options)
tk.Button(add_frame, text="Вставить", command=lambda: paste_from_clipboard(source_entry)).grid(row=0, column=3, **pad_options)
tk.Button(add_frame, text="Удалить", command=lambda: clear_entry(source_entry)).grid(row=0, column=4, **pad_options)

# Поле целевой папки
tk.Label(add_frame, text="Целевая папка:", font=font_bold).grid(row=1, column=0, **pad_options)
target_entry = tk.Entry(add_frame, width=80, font=font_normal)
target_entry.grid(row=1, column=1, **pad_options)
tk.Button(add_frame, text="Обзор", command=lambda: browse_folder(target_entry)).grid(row=1, column=2, **pad_options)
tk.Button(add_frame, text="Вставить", command=lambda: paste_from_clipboard(target_entry)).grid(row=1, column=3, **pad_options)
tk.Button(add_frame, text="Удалить", command=lambda: clear_entry(target_entry)).grid(row=1, column=4, **pad_options)

# Поле имени симлинка
tk.Label(add_frame, text="Имя симлинка:", font=font_bold).grid(row=2, column=0, **pad_options)
name_entry = tk.Entry(add_frame, width=80, font=font_normal)
name_entry.grid(row=2, column=1, **pad_options)
tk.Button(add_frame, text="Вставить", command=lambda: paste_from_clipboard(name_entry)).grid(row=2, column=3, **pad_options)
tk.Button(add_frame, text="Удалить", command=lambda: clear_entry(name_entry)).grid(row=2, column=4, **pad_options)

# Чекбокс для ручного задания имени симлинка
manual_name_var = tk.BooleanVar(value=False)
tk.Checkbutton(add_frame, text="Задать имя вручную", variable=manual_name_var, font=font_normal,
               command=update_symlink_name).grid(row=3, column=1, sticky='w', pady=5)

# Кнопка создания
tk.Button(add_frame, text="СОЗДАТЬ СИМЛИНК", command=create_symlink,
         font=('Arial', 12, 'bold'), bg='#4CAF50', fg='white').grid(row=4, column=0, columnspan=5, pady=20)

# Вкладка "Поиск и добавление"
search_frame = ttk.Frame(notebook)
notebook.add(search_frame, text="Поиск и добавление")

# Чекбоксы для типов файлов
tk.Label(search_frame, text="Искать файлы:", font=font_bold).grid(row=0, column=0, **pad_options)
dll_var = tk.BooleanVar(value=True)
vst3_var = tk.BooleanVar(value=True)
tk.Checkbutton(search_frame, text="*.dll", variable=dll_var, font=font_normal).grid(row=0, column=1, sticky='w')
tk.Checkbutton(search_frame, text="*.vst3", variable=vst3_var, font=font_normal).grid(row=0, column=2, sticky='w')

# Поле выбора диска
tk.Label(search_frame, text="Выберите диск:", font=font_bold).grid(row=1, column=0, **pad_options)
drive_var = tk.StringVar()
drive_menu = ttk.Combobox(search_frame, textvariable=drive_var, values=get_available_drives(), state="readonly")
drive_menu.grid(row=1, column=1, **pad_options)
search_button = tk.Button(search_frame, text="Поиск", command=search_files)
search_button.grid(row=1, column=2, **pad_options)
cancel_button = tk.Button(search_frame, text="Отмена", command=cancel_search, state="disabled")
cancel_button.grid(row=1, column=3, **pad_options)

# Элементы управления выбором
tk.Checkbutton(search_frame, text="Выбрать все", command=select_all, font=font_normal).grid(row=2, column=0, sticky='w', pady=5)
tk.Button(search_frame, text="Отменить выбор", command=deselect_all).grid(row=2, column=1, sticky='w', pady=5)

# Таблица для отображения файлов
tree = ttk.Treeview(search_frame, columns=("Check", "Path", "ParentFolder"), show="headings", selectmode="extended")
tree.heading("Check", text="Выбрать")
tree.heading("Path", text="Путь к файлу")
tree.heading("ParentFolder", text="Родительская папка")
tree.column("Check", width=60, anchor="center")
tree.column("Path", width=500)
tree.column("ParentFolder", width=200)
tree.grid(row=3, column=0, columnspan=5, sticky="nsew", padx=5, pady=5)
scrollbar = ttk.Scrollbar(search_frame, orient="vertical", command=tree.yview)
scrollbar.grid(row=3, column=5, sticky="ns")
tree.configure(yscrollcommand=scrollbar.set)
tree.bind("<Button-1>", toggle_check)

# Метка прогресса
progress_label = tk.Label(search_frame, text="", font=font_normal)
progress_label.grid(row=4, column=0, columnspan=5, pady=5)

# Поле целевой папки
tk.Label(search_frame, text="Целевая папка:", font=font_bold).grid(row=5, column=0, **pad_options)
search_target_entry = tk.Entry(search_frame, width=80, font=font_normal)
search_target_entry.grid(row=5, column=1, **pad_options)
tk.Button(search_frame, text="Обзор", command=lambda: browse_folder(search_target_entry)).grid(row=5, column=2, **pad_options)
tk.Button(search_frame, text="Вставить", command=lambda: paste_from_clipboard(search_target_entry)).grid(row=5, column=3, **pad_options)
tk.Button(search_frame, text="Удалить", command=lambda: clear_entry(search_target_entry)).grid(row=5, column=4, **pad_options)

# Кнопка создания симлинков
tk.Button(search_frame, text="СОЗДАТЬ СИМЛИНКИ", command=create_selected_symlinks,
         font=('Arial', 12, 'bold'), bg='#4CAF50', fg='white').grid(row=6, column=0, columnspan=5, pady=20)

# Глобальные переменные
found_files = []
check_vars = []
search_cancelled = False

# Горячие клавиши
root.bind('<Control-v>', lambda e: paste_from_clipboard(root.focus_get()))
root.bind('<Button-3>', lambda e: paste_from_clipboard(root.focus_get()))

# Отслеживание изменений в поле исходной папки
source_entry.bind('<KeyRelease>', update_symlink_name)

if not is_admin():
    messagebox.showwarning("Внимание", "Для создания симлинков требуются права администратора")

root.mainloop()