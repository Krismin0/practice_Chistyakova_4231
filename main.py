import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from generate_file1 import generate_conference_program
from generate_file2 import generate_conference_report
from generate_file3 import create_accepted_papers_list
import json
import os


# Стиль для всего приложения
def configure_styles():
    style = ttk.Style()
    style.configure('TFrame', background='#f0f0f0')
    style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
    style.configure('TButton', font=('Arial', 10), padding=5)
    style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
    style.configure('Accent.TButton', font=('Arial', 10, 'bold'), foreground='black')
    style.map('Accent.TButton',
              background=[('active', '#45a049'), ('!disabled', '#4CAF50')])


def open_data_window(root, report_mode=False, papers_mode=False):
    root.withdraw()  # Скрыть главное окно

    def check_fields():
        if papers_mode:
            required_fields = ["conf_number", "leader_name", "leader_email", "leader_phone"]
        else:
            required_fields = ["conf_number", "head", "secretary"] if report_mode else list(entries.keys())

        all_filled = all(entries[field].get().strip() for field in required_fields) and json_path.get()
        create_btn.config(state=tk.NORMAL if all_filled else tk.DISABLED)

    def select_json_file():
        filename = filedialog.askopenfilename(
            title="Выберите файл с данными участников",
            filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")]
        )
        if filename:
            json_path.set(filename)
            check_fields()

    def create_file():
        try:
            if papers_mode:
                conf_number = entries["conf_number"].get().strip()
                default_filename = f"Список представляемых к публикации докладов конференции {conf_number}.docx" if conf_number else "Список представляемых к публикации докладов.docx"

                output_path = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word документ", "*.docx")],
                    title="Сохранить список публикуемых докладов",
                    initialfile=default_filename
                )
                if not output_path:
                    return

                create_accepted_papers_list(
                    json_path.get(),
                    output_path,
                    entries["leader_name"].get(),
                    entries["leader_email"].get(),
                    entries["leader_phone"].get(),
                    conf_number
                )
            else:
                conference_data = {
                    "number": entries["conf_number"].get(),
                    "head": entries["head"].get(),
                    "deputy": entries.get("deputy", tk.StringVar(value="")).get(),
                    "secretary": entries.get("secretary", tk.StringVar(value="")).get()
                }

                default_filename = f"{'Отчет' if report_mode else 'Программа'}_конференции_{conference_data['number']}.docx"
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word документ", "*.docx")],
                    title="Сохранить документ конференции",
                    initialfile=default_filename
                )
                if not output_path:
                    return

                with open(json_path.get(), encoding="utf-8") as f:
                    contributions = json.load(f)

                if report_mode:
                    generate_conference_report(conference_data, contributions, output_path)
                else:
                    generate_conference_program(conference_data, contributions, output_path)

            messagebox.showinfo("Успех", "Документ успешно создан!")
            window.destroy()
            root.deiconify()  # Показать главное окно снова

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать файл:\n{e}")

    window = tk.Toplevel()
    window.protocol("WM_DELETE_WINDOW", lambda: (window.destroy(), root.deiconify()))  # При закрытии вернуть главное окно

    window.title("Данные для списка публикуемых докладов" if papers_mode else (
        "Данные для отчета" if report_mode else "Данные для программы конференции"))
    window.configure(bg='#f0f0f0')
    window.resizable(False, False)

    title_text = "Данные для списка публикуемых докладов" if papers_mode else (
        "Данные для отчёта о конференции" if report_mode else "Данные для программы конференции")

    ttk.Label(window, text=title_text, style='Header.TLabel').grid(
        row=0, column=0, columnspan=3, pady=(10, 20), padx=10, sticky="w")

    if papers_mode:
        fields = [
            ("Номер конференции (например, 77-й):", "conf_number"),
            ("ФИО руководителя УНИДС:", "leader_name"),
            ("Email руководителя УНИДС:", "leader_email"),
            ("Телефон руководителя УНИДС:", "leader_phone")
        ]
    elif report_mode:
        fields = [
            ("Номер конференции (например, 77-й):", "conf_number"),
            ("Научный руководитель (ФИО, должность, звание):", "head"),
            ("Секретарь (ФИО, должность):", "secretary")
        ]
    else:
        fields = [
            ("Номер конференции (например, 77-й):", "conf_number"),
            ("Научный руководитель (ФИО, должность, звание):", "head"),
            ("Заместитель (ФИО, должность, звание):", "deputy"),
            ("Секретарь (ФИО, должность):", "secretary")
        ]

    entries = {}

    def bind_paste_shortcuts(entry):
        def handle_keys(event):
            # Ctrl+V → вставка
            if (event.state & 0x4) and event.keycode == 86:
                try:
                    content = event.widget.clipboard_get()
                    event.widget.insert(tk.INSERT, content)
                except tk.TclError:
                    pass
                return "break"

            # Ctrl+A → выделить всё
            elif (event.state & 0x4) and event.keycode == 65:
                event.widget.select_range(0, 'end')
                event.widget.icursor('end')
                return "break"

            # Ctrl+C → копировать
            elif (event.state & 0x4) and event.keycode == 67:
                try:
                    selection = event.widget.selection_get()
                    event.widget.clipboard_clear()
                    event.widget.clipboard_append(selection)
                except tk.TclError:
                    pass
                return "break"

        entry.bind("<KeyPress>", handle_keys)

    for i, (label_text, field_name) in enumerate(fields, start=1):
        ttk.Label(window, text=label_text).grid(row=i, column=0, sticky="w", padx=10, pady=5)
        entry = ttk.Entry(window, width=40)
        entry.grid(row=i, column=1, padx=10, pady=5, columnspan=2, sticky="ew")
        bind_paste_shortcuts(entry)
        entry.bind("<KeyRelease>", lambda e: check_fields())
        entries[field_name] = entry

    row_num = len(fields) + 1
    ttk.Label(window, text="Файл с данными участников:").grid(row=row_num, column=0, sticky="w", padx=10, pady=5)

    json_path = tk.StringVar()
    entry_file = ttk.Entry(window, textvariable=json_path, width=30, state='readonly')
    entry_file.grid(row=row_num, column=1, padx=10, pady=5, sticky="ew")

    btn_file = ttk.Button(window, text="Выбрать...", command=select_json_file)
    btn_file.grid(row=row_num, column=2, padx=(0, 10), pady=5, sticky="e")

    btn_text = "Создать список публикуемых докладов" if papers_mode else (
        "Создать отчёт" if report_mode else "Создать программу конференции")

    create_btn = ttk.Button(
        window, text=btn_text, command=create_file,
        style='Accent.TButton', state=tk.DISABLED
    )
    create_btn.grid(row=row_num + 1, column=0, columnspan=3, pady=(20, 10), padx=10, sticky="ew")

    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'+{x}+{y}')


def run_interface():
    root = tk.Tk()
    root.title("Генератор документов конференции")
    root.configure(bg='#f0f0f0')
    root.resizable(False, False)

    configure_styles()

    width = 400
    height = 300
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    main_frame = ttk.Frame(root)
    main_frame.pack(expand=True, fill='both', padx=20, pady=20)

    ttk.Label(
        main_frame,
        text="Генератор документов конференции",
        style='Header.TLabel'
    ).pack(pady=(0, 20))

    ttk.Button(
        main_frame,
        text="Создать программу конференции",
        command=lambda: open_data_window(root, report_mode=False),
        style='Accent.TButton'
    ).pack(fill='x', pady=10, ipady=8)

    ttk.Button(
        main_frame,
        text="Создать отчёт о конференции",
        command=lambda: open_data_window(root, report_mode=True),
        style='Accent.TButton'
    ).pack(fill='x', pady=10, ipady=8)

    ttk.Button(
        main_frame,
        text="Создать список публикуемых докладов",
        command=lambda: open_data_window(root, papers_mode=True),
        style='Accent.TButton'
    ).pack(fill='x', pady=10, ipady=8)

    root.mainloop()


if __name__ == "__main__":
    run_interface()
