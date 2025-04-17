import io
import fitz  # PyMuPDF
import pandas as pd
from tkinter import Tk, Canvas, Scrollbar, filedialog, Button, Label, Entry, Frame, Scale, OptionMenu, StringVar, \
    messagebox
from PIL import Image, ImageTk

# Глобальные переменные для хранения данных и параметров
excel_data = None  # Словарь данных из Excel для выбранного листа
global_horiz_offset = 0.0  # Горизонтальное смещение для вставки текста
global_vert_offset = 0.0  # Вертикальное смещение для вставки текста
current_scale = 100  # Масштаб предпросмотра (в процентах)
pdf_page_count = 1  # Число страниц в выбранном PDF


def select_excel_file():
    """Выбор Excel-файла и заполнение выпадающего списка доступных листов."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    if file_path:
        excel_label.config(text=file_path)
        try:
            xls = pd.ExcelFile(file_path)
            sheets = xls.sheet_names
            # Задаём первым выбранный лист
            sheet_var.set(sheets[0])
            sheet_menu['menu'].delete(0, 'end')
            for sheet in sheets:
                sheet_menu['menu'].add_command(label=sheet, command=lambda value=sheet: sheet_var.set(value))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка чтения Excel файла:\n{e}")
    return file_path


def select_pdf_file():
    """Выбор PDF-файла и получение числа его страниц."""
    global pdf_page_count
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_label.config(text=file_path)
        try:
            doc = fitz.open(file_path)
            pdf_page_count = doc.page_count
            doc.close()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка открытия PDF:\n{e}")
    return file_path


def select_output_file():
    """Выбор пути для сохранения итогового PDF. После выбора сразу запускается сохранение."""
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF files", "*.pdf")])
    if file_path:
        output_entry.delete(0, 'end')
        output_entry.insert(0, file_path)
        transfer_data()
    return file_path


def refresh_preview():
    """
    Открывает выбранный PDF, для каждой страницы:
      – накладывает данные из Excel (если они применены) с учётом смещений,
      – рендерит страницу с выбранным масштабом,
      – располагает полученные изображения подряд по вертикали.
    Итоговое изображение выводится в scrollable Canvas.
    """
    pdf_file = pdf_label.cget("text")
    if not pdf_file or pdf_file == "Файл не выбран":
        return
    try:
        doc = fitz.open(pdf_file)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка открытия PDF:\n{e}")
        return

    # Очищаем Canvas и список ссылок для изображений
    preview_canvas.delete("all")
    preview_canvas.image_refs = []

    scale_factor = current_scale / 100.0
    mat = fitz.Matrix(scale_factor, scale_factor)

    y_position = 0  # Начальное смещение для размещения страниц
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        # Если Excel-данные уже применены, накладываем данные
        if excel_data is not None:
            blocks = page.get_text("blocks")
            for block in blocks:
                key = block[4].strip().lower()
                if key in excel_data:
                    value = excel_data[key]
                    text_insert = f"{value:.3f}"
                    x0, y0, x1, y1 = block[:4]
                    x_insert = x1 + global_horiz_offset
                    y_insert = y1 - global_vert_offset
                    page.insert_text((x_insert, y_insert), text_insert, fontname="Times-Roman", fontsize=8)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("ppm")
        image = Image.open(io.BytesIO(img_data))
        photo = ImageTk.PhotoImage(image)
        preview_canvas.create_image(0, y_position, anchor="nw", image=photo)
        preview_canvas.image_refs.append(photo)  # сохраняем ссылку, чтобы изображение не удалялось
        y_position += photo.height()
    doc.close()
    preview_canvas.config(scrollregion=(0, 0, preview_canvas.winfo_width(), y_position))


def on_scale_change(new_value):
    """Обработчик изменения значения ползунка масштаба – обновление предпросмотра."""
    global current_scale
    current_scale = int(float(new_value))
    refresh_preview()


def apply_excel_data():
    """
    Считывает данные из выбранного листа Excel и значения смещений,
    затем обновляет предпросмотр PDF.
    """
    global global_horiz_offset, global_vert_offset, excel_data
    excel_file = excel_label.cget("text")
    if not excel_file or excel_file == "Файл не выбран":
        messagebox.showerror("Ошибка", "Выберите Excel-файл")
        return
    # Получаем выбранный лист
    sheet = sheet_var.get()
    if not sheet:
        messagebox.showerror("Ошибка", "Выберите лист Excel")
        return
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        df["Цинк (Zn). %"] = pd.to_numeric(df["Цинк (Zn). %"], errors='coerce')
        excel_data = df.set_index("номер пробы", drop=True)["Цинк (Zn). %"] \
            .apply(lambda x: round(x, 3)).to_dict()
        # Приводим ключи к строковому и нижнему регистру
        excel_data = {str(k).lower(): v for k, v in excel_data.items()}
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка обработки Excel:\n{e}")
        return
    try:
        global_horiz_offset = float(preview_horiz_offset_entry.get())
        global_vert_offset = float(preview_vert_offset_entry.get())
    except Exception as e:
        messagebox.showerror("Ошибка", f"Неверное значение смещения:\n{e}")
        return
    refresh_preview()


def transfer_data():
    """
    Обрабатывает весь PDF – накладывает данные из Excel (с учётом смещений)
    на все страницы и сохраняет итоговый файл по выбранному пути.
    """
    if excel_data is None:
        messagebox.showerror("Ошибка", "Сначала нажмите «Примерить данные» для применения Excel-данных")
        return
    excel_file = excel_label.cget("text")
    pdf_file = pdf_label.cget("text")
    output_file = output_entry.get()
    if not excel_file or not pdf_file or not output_file:
        messagebox.showerror("Ошибка", "Выберите Excel, PDF и укажите путь для сохранения")
        return
    try:
        doc = fitz.open(pdf_file)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка открытия PDF:\n{e}")
        return

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        blocks = page.get_text("blocks")
        for block in blocks:
            key = block[4].strip().lower()
            if key in excel_data:
                value = excel_data[key]
                text_insert = f"{value:.3f}"
                x0, y0, x1, y1 = block[:4]
                x_insert = x1 + global_horiz_offset
                y_insert = y1 - global_vert_offset
                page.insert_text((x_insert, y_insert), text_insert, fontname="Times-Roman", fontsize=8)
    try:
        doc.save(output_file)
        doc.close()
        messagebox.showinfo("Успех", f"PDF успешно сохранён:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка сохранения PDF:\n{e}")
        doc.close()


def _on_mousewheel(event):
    """Обработчик прокрутки колесиком мыши для Canvas."""
    preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


# ------------------- Интерфейс -------------------
root = Tk()
root.title("Инструмент переноса данных и предпросмотра PDF")
root.geometry("1400x1200")
root.configure(bg="#f0f0f0")

# Левая панель: выбор файлов, листа и путь для сохранения
controls_frame = Frame(root, bg="#f0f0f0")
controls_frame.pack(side="left", fill="y", padx=10, pady=10)

Label(controls_frame, text="Выберите Excel-файл", bg="#f0f0f0").pack(pady=5)
Button(controls_frame, text="Обзор", command=select_excel_file).pack(pady=5)
excel_label = Label(controls_frame, text="Файл не выбран", bg="#f0f0f0", wraplength=300)
excel_label.pack(pady=5)

Label(controls_frame, text="Выберите лист Excel", bg="#f0f0f0").pack(pady=5)
sheet_var = StringVar()
sheet_menu = OptionMenu(controls_frame, sheet_var, "")
sheet_menu.pack(pady=5)

Label(controls_frame, text="Выберите PDF-файл", bg="#f0f0f0").pack(pady=5)
Button(controls_frame, text="Обзор", command=select_pdf_file).pack(pady=5)
pdf_label = Label(controls_frame, text="Файл не выбран", bg="#f0f0f0", wraplength=300)
pdf_label.pack(pady=5)

Label(controls_frame, text="Путь для сохранения PDF", bg="#f0f0f0").pack(pady=5)
output_entry = Entry(controls_frame, width=40)
output_entry.pack(pady=5)
Button(controls_frame, text="Сохранить как", command=select_output_file).pack(pady=20)

# Правая панель: предпросмотр PDF с прокруткой
preview_frame = Frame(root, bg="#ffffff")
preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

preview_container = Frame(preview_frame, bg="#ffffff")
preview_container.pack(fill="both", expand=True)

# Верхний блок управления предпросмотром: масштаб и смещения
preview_controls_frame = Frame(preview_container, bg="#e0e0e0", bd=2, relief="groove")
preview_controls_frame.pack(side="top", fill="x", padx=5, pady=5)

Label(preview_controls_frame, text="Масштаб:", bg="#e0e0e0") \
    .grid(row=0, column=0, padx=5, pady=5, sticky="w")
scale_slider = Scale(preview_controls_frame, from_=50, to=200, orient="horizontal",
                     resolution=5, command=on_scale_change, bg="#e0e0e0")
scale_slider.set(current_scale)
scale_slider.grid(row=0, column=1, padx=5, pady=5, sticky="we")

Label(preview_controls_frame, text="Гориз. смещение:", bg="#e0e0e0") \
    .grid(row=1, column=0, padx=5, pady=5, sticky="w")
preview_horiz_offset_entry = Entry(preview_controls_frame, width=8)
preview_horiz_offset_entry.insert(0, "180")
preview_horiz_offset_entry.grid(row=1, column=1, padx=5, pady=5)

Label(preview_controls_frame, text="Вер. смещение:", bg="#e0e0e0") \
    .grid(row=2, column=0, padx=5, pady=5, sticky="w")
preview_vert_offset_entry = Entry(preview_controls_frame, width=8)
preview_vert_offset_entry.insert(0, "1.2")
preview_vert_offset_entry.grid(row=2, column=1, padx=5, pady=5)

Button(preview_controls_frame, text="Примерить данные", command=apply_excel_data) \
    .grid(row=3, column=0, columnspan=2, padx=5, pady=5)

# Блок для Canvas с вертикальным скроллом
canvas_frame = Frame(preview_container, bg="#ffffff")
canvas_frame.pack(side="top", fill="both", expand=True)

scrollbar = Scrollbar(canvas_frame, orient="vertical")
scrollbar.pack(side="right", fill="y")

preview_canvas = Canvas(canvas_frame, bg="#cccccc", yscrollcommand=scrollbar.set)
preview_canvas.pack(side="left", fill="both", expand=True)
scrollbar.config(command=preview_canvas.yview)

# Привязка колесика мыши для прокрутки Canvas
preview_canvas.bind("<Enter>", lambda event: preview_canvas.bind_all("<MouseWheel>", _on_mousewheel))
preview_canvas.bind("<Leave>", lambda event: preview_canvas.unbind_all("<MouseWheel>"))

root.mainloop()
