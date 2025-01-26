import os
import time
import win32print
from tkinter import Tk, filedialog, Button, Label

def print_files():
    # Открываем диалог выбора папки
    folder_path = filedialog.askdirectory(title="Выберите папку с файлами для печати")
    
    if folder_path:
        # Получаем список всех файлов в папке
        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

        if not files:
            label.config(text="В папке нет файлов для печати.")
            return

        try:
            # Получаем принтер по умолчанию
            printer_name = win32print.GetDefaultPrinter()
            label.config(text=f"Принтер: {printer_name}")

            # Отправляем файлы на печать
            for file_path in files:
                try:
                    os.startfile(file_path, "print")
                    label.config(text=f"Печать: {os.path.basename(file_path)}")
                    time.sleep(2)  # Задержка между заданиями
                except Exception as e:
                    label.config(text=f"Ошибка при печати файла {os.path.basename(file_path)}: {e}")

            label.config(text="Все файлы отправлены на печать.")
        except Exception as e:
            label.config(text=f"Ошибка: {e}")
    else:
        label.config(text="Папка не выбрана.")

# Создаем главное окно приложения
app = Tk()
app.title("Пакетная печать документов")
app.geometry("400x200")

# Кнопка для запуска диалога выбора папки и печати файлов
button = Button(app, text="Выбрать папку и печатать", command=print_files)
button.pack(pady=20)

# Метка для отображения статуса
label = Label(app, text="")
label.pack()

# Запуск главного цикла приложения
app.mainloop()
