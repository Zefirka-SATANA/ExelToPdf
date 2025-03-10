import pandas as pd
import fitz
from tkinter import Tk, filedialog, Button, Label, Entry, messagebox, OptionMenu, StringVar, Frame

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    excel_label.config(text=file_path)

    # Получаем все листы из выбранного файла Excel
    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names

    # Очистка предыдущих опций и добавление новых
    sheet_var.set('')
    sheet_menu['menu'].delete(0, 'end')
    for sheet in sheets:
        sheet_menu['menu'].add_command(label=sheet, command=lambda value=sheet: sheet_var.set(value))

    return file_path

def select_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_label.config(text=file_path)
    return file_path

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    output_entry.delete(0, 'end')
    output_entry.insert(0, file_path)
    return file_path

def transfer_data():
    excel_file = excel_label.cget("text")
    pdf_file = pdf_label.cget("text")
    output_file = output_entry.get()
    sheet_name = sheet_var.get()
    horizontal_offset = float(horizontal_offset_entry.get())
    vertical_offset = float(vertical_offset_entry.get())

    # Загружаем данные из выбранного листа Excel
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df["Цинк (Zn). %"] = pd.to_numeric(df["Цинк (Zn). %"], errors='coerce')  # Преобразуем в числовой формат
    sample_data = df.set_index("номер пробы", drop=True)["Цинк (Zn). %"].apply(lambda x: round(x, 3)).to_dict()
    sample_data = {k.lower(): v for k, v in sample_data.items()}  # Преобразуем ключи к нижнему регистру

    # Загружаем PDF файл
    pdf_document = fitz.open(pdf_file)
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        blocks = page.get_text("blocks")

        # Обрабатываем блоки текста и вставляем данные в нужный столбец
        for block in blocks:
            key = block[4].strip().lower()  # Преобразуем ключ к нижнему регистру
            if key in sample_data:
                value = sample_data[key]
                text_insert = f"{value:.3f}"

                # Координаты блока (x0, y0, x1, y1)
                x0, y0, x1, y1 = block[:4]
                x_insert = x1 + horizontal_offset  # Смещаем текст вправо от блока
                y_insert = y1 - vertical_offset  # Смещаем текст вертикально

                page.insert_text((x_insert, y_insert), text_insert, fontname="Times-Roman", fontsize=8)

    pdf_document.save(output_file)
    pdf_document.close()

    messagebox.showinfo("Успех", f"Данные успешно перенесены в {output_file}")

# Создаем основное окно
root = Tk()
root.title("Инструмент переноса данных")
root.geometry("450x600")
root.configure(bg="#f0f0f0")

# Создаем контейнер для элементов
frame = Frame(root, bg="#f0f0f0")
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Элементы UI
excel_label = Label(frame, text="Выберите файл Excel", bg="#f0f0f0")
excel_label.pack(pady=5)
excel_button = Button(frame, text="Обзор", command=select_excel_file)
excel_button.pack(pady=5)

sheet_label = Label(frame, text="Выберите лист", bg="#f0f0f0")
sheet_label.pack(pady=5)
sheet_var = StringVar()
sheet_menu = OptionMenu(frame, sheet_var, '')
sheet_menu.pack(pady=5)

pdf_label = Label(frame, text="Выберите файл PDF", bg="#f0f0f0")
pdf_label.pack(pady=5)
pdf_button = Button(frame, text="Обзор", command=select_pdf_file)
pdf_button.pack(pady=5)

output_label = Label(frame, text="Путь для сохранения", bg="#f0f0f0")
output_label.pack(pady=5)
output_entry = Entry(frame)
output_entry.pack(pady=5)
output_button = Button(frame, text="Сохранить как", command=select_output_file)
output_button.pack(pady=5)

# Создаем рамку для ввода смещений
offset_frame = Frame(frame, bg="#d3d3d3", bd=2, relief="groove")
offset_frame.pack(pady=10, padx=10, fill="x", expand=True)

offset_label = Label(offset_frame, text="Трогать если вставилось криво:", bg="#d3d3d3")
offset_label.pack(pady=5)

horizontal_label = Label(offset_frame, text="Горизонталь", bg="#d3d3d3")
horizontal_label.pack(pady=5)
horizontal_offset_entry = Entry(offset_frame)
horizontal_offset_entry.insert(0, "180")  # Значение по умолчанию
horizontal_offset_entry.pack(pady=5)

vertical_label = Label(offset_frame, text="Вертикаль", bg="#d3d3d3")
vertical_label.pack(pady=5)
vertical_offset_entry = Entry(offset_frame)
vertical_offset_entry.insert(0, "1.2")  # Значение по умолчанию
vertical_offset_entry.pack(pady=5)

transfer_button = Button(frame, text="Перенести данные", command=transfer_data)
transfer_button.pack(pady=20)

root.mainloop()
