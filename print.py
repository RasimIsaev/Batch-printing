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

        # GUI Elements
        self.setup_ui()

    def setup_ui(self):
        # Frame for buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        # Add File Button
        self.add_btn = ttk.Button(button_frame, text="Добавить файлы", command=self.add_file)
        self.add_btn.pack(side=tk.LEFT, padx=5)

        # Remove File Button
        self.remove_btn = ttk.Button(button_frame, text="Удалить файлы", command=self.remove_file)
        self.remove_btn.pack(side=tk.LEFT, padx=5)

        # Printer Selection
        self.printer_var = tk.StringVar()
        self.printer_combobox = ttk.Combobox(
            self.root, 
            textvariable=self.printer_var, 
            values=self.get_available_printers(),
            state="readonly"
        )
        self.printer_combobox.set(self.printer_name)
        self.printer_combobox.pack(pady=5)

        # Spinbox for selecting number of copies
        self.copies_var = tk.IntVar(value=1)
        self.copies_spinbox = ttk.Spinbox(
            self.root, from_=1, to=100, textvariable=self.copies_var, width=5
        )
        self.copies_spinbox.pack(pady=5)

        # Listbox for files
        self.listbox = tk.Listbox(self.root, selectmode=tk.EXTENDED, width=50)
        self.listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Print Button
        self.print_btn = ttk.Button(self.root, text="Печать", command=self.print_files)
        self.print_btn.pack(pady=10)

    def get_available_printers(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        return [printer[2] for printer in printers]

    def add_file(self):
        files = filedialog.askopenfilenames(
            title="Select Files",
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

    def print_files(self):
        if not self.files_list:
            messagebox.showerror("Ошибка", "Файлы не выбран\u044b!")
            return

        printer = self.printer_var.get()
        if not printer:
            messagebox.showerror("Ошибка", "Принтер не выб\u0440\u0430\u043d!")
            return

        # Get the number of copies
        copies = self.copies_var.get()

        # Confirmation
        if not messagebox.askyesno("Подтверж\u0434\u0435\u043d\u0438\u0435", f"Отправи\u0442\u044c указан\u043d\u044b\u0435 фай\u043b\u044b на печат\u044c {printer}?"):
            return

        # Print files with delay between copies
        for file in self.files_list:
            try:
                for _ in range(copies):
                    win32api.ShellExecute(0, "print", file, f'"{printer}"', ".", 0)
                    time.sleep(2)  # Добавляем задержку в 2 секунды
            except Exception as e:
                messagebox.showerror("Print Error", f"Failed to print {file}: {str(e)}")

        messagebox.showinfo("Ув\u0435дом\u043bен\u0438\u0435", "З\u0430д\u0430н\u0438\u044f н\u0430 печат\u044c о\u0442п\u0440а\u0432\u043b\u0435\u043d\u044b!")

if __name__ == "__main__":
    root = tk.Tk()
    app = BatchPrinterApp(root)
    root.mainloop()
