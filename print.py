import os
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32print
import win32api

class BatchPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Массовая печать")
        self.files_list = []
        self.printer_name = win32print.GetDefaultPrinter()  # Получить принтер по умолчанию

        # Элементы графического интерфейса (GUI)
        self.setup_ui()

    def setup_ui(self):
        # Фрейм для кнопок
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        # Кнопка добавления файла
        self.add_btn = ttk.Button(button_frame, text="Добавить файлы", command=self.add_file)
        self.add_btn.pack(side=tk.LEFT, padx=5)

        # Кнопка удаления файла
        self.remove_btn = ttk.Button(button_frame, text="Удалить файлы", command=self.remove_file)
        self.remove_btn.pack(side=tk.LEFT, padx=5)

        # Выбор принтера
        self.printer_var = tk.StringVar()
        self.printer_combobox = ttk.Combobox(
            self.root, 
            textvariable=self.printer_var, 
            values=self.get_available_printers(),
            state="readonly"
        )
        self.printer_combobox.set(self.printer_name)
        self.printer_combobox.pack(pady=5)

        # Счетчик (Spinbox) для выбора количества копий
        self.copies_var = tk.IntVar(value=1)
        self.copies_spinbox = ttk.Spinbox(
            self.root, from_=1, to=100, textvariable=self.copies_var, width=5
        )
        self.copies_spinbox.pack(pady=5)

        # Список (Listbox) для файлов с функцией перетаскивания (drag-and-drop)
        self.listbox = tk.Listbox(self.root, selectmode=tk.EXTENDED, width=50)
        self.listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.listbox.bind("<Button-1>", self.start_drag)
        self.listbox.bind("<B1-Motion>", self.do_drag)
        self.listbox.bind("<ButtonRelease-1>", self.end_drag)
        self.dragging_index = None

        # Кнопка печати
        self.print_btn = ttk.Button(self.root, text="Печать", command=self.print_files)
        self.print_btn.pack(pady=10)

    def get_available_printers(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        return [printer[2] for printer in printers]

    def add_file(self):
        files = filedialog.askopenfilenames(
            title="Выбрать файлы",
            filetypes=[("\u0412\u0441\u0435 \u0444\u0430\u0439\u043b\u044b", "*.*"), ("PDF Files", "*.pdf"), ("Text Files", "*.txt")]
        )
        for file in files:
            if file not in self.files_list:
                self.files_list.append(file)
                self.listbox.insert(tk.END, os.path.basename(file))

    def remove_file(self):
        selected = self.listbox.curselection()
        for index in reversed(selected):
            self.listbox.delete(index)
            del self.files_list[index]

    def start_drag(self, event):
        self.dragging_index = self.listbox.nearest(event.y)

    def do_drag(self, event):
        current_index = self.listbox.nearest(event.y)
        if current_index != self.dragging_index and self.dragging_index is not None:
            # Обмен местами элементов в списке и во внутреннем списке
            self.files_list[self.dragging_index], self.files_list[current_index] = (
                self.files_list[current_index], self.files_list[self.dragging_index]
            )
            item_text = self.listbox.get(self.dragging_index)
            self.listbox.delete(self.dragging_index)
            self.listbox.insert(current_index, item_text)
            self.listbox.selection_set(current_index)
            self.dragging_index = current_index

    def end_drag(self, event):
        self.dragging_index = None

    def print_files(self):
        if not self.files_list:
            messagebox.showerror("Ошибка", "Файл\u044b не в\u044b\u0431р\u0430\u043d\u044b!")
            return

        printer = self.printer_var.get()
        if not printer:
            messagebox.showerror("Ошиб\u043a\u0430", "П\u0440\u0438\u043d\u0442\u0435\u0440 н\u0435 в\u044b\u0431\u0440\u0430\u043d!")
            return

        # Получение количества копий
        copies = self.copies_var.get()

        # Подтверждение
        if not messagebox.askyesno("П\u043eд\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043d\u0438\u0435", f"Отп\u0440\u0430\u0432\u0438\u0442\u044c у\u043a\u0430\u0437\u0430\u043d\u043d\u044b\u0435 ф\u0430й\u043b\u044b н\u0430 \u043f\u0435\u0447ат\u044c {printer}?"):
            return

        # Печать файлов с задержкой между копиями
        for file in self.files_list:
            try:
                for _ in range(copies):
                    win32api.ShellExecute(0, "print", file, f'"{printer}"', ".", 0)
                    time.sleep(2)  # \u0414\u043e\u0431\u0430\u0432\u043b\u044f\u0435\u043c за\u0434\u0435\u0440\u0436\u043a\u0443 в 2 с\u0435\u043a\u0443\u043d\u0434\u044b
            except Exception as e:
                messagebox.showerror("Print Error", f"Failed to print {file}: {str(e)}")

        messagebox.showinfo("У\u0432е\u0434\u043e\u043c\u043b\u0435\u043d\u0438\u0435", "\u0417\u0430\u0434\u0430\u043d\u0438\u044f \u043d\u0430 \u043f\u0435\u0447\u0430\u0442\u044c \u043e\u0442\u043f\u0440\u0430\u0432\u043b\u0435\u043d\u044b!")

if __name__ == "__main__":
    root = tk.Tk()
    app = BatchPrinterApp(root)
    root.mainloop()
