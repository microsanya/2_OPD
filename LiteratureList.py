import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import requests

class ReferenceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Добавление источников в Word-документ")
        self.root.geometry("800x600")  # Установка размера окна

        # Тип источника
        self.source_type_label = tk.Label(root, text="Тип источника:")
        self.source_type_label.pack()
        self.source_type_var = tk.StringVar()
        self.source_type_entry = tk.Entry(root, textvariable=self.source_type_var)
        self.source_type_entry.pack()

        # Автор
        self.author_label = tk.Label(root, text="Автор:")
        self.author_label.pack()
        self.author_var = tk.StringVar()
        self.author_entry = tk.Entry(root, textvariable=self.author_var)
        self.author_entry.pack()

        # Название
        self.title_label = tk.Label(root, text="Название:")
        self.title_label.pack()
        self.title_var = tk.StringVar()
        self.title_entry = tk.Entry(root, textvariable=self.title_var)
        self.title_entry.pack()

        # Выбор пути для сохранения файла
        self.save_path_button = tk.Button(root, text="Выбрать путь для сохранения файла", command=self.select_save_path)
        self.save_path_button.pack()

        # Сохранить
        self.save_button = tk.Button(root, text="Сохранить", command=self.save_references)
        self.save_button.pack()

        self.save_path = None
        self.references = []

    def select_save_path(self):
        self.save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if self.save_path:
            messagebox.showinfo("Путь для сохранения выбран", f"Файл будет сохранен по пути: {self.save_path}")

    def save_references(self):
        if not self.save_path:
            messagebox.showerror("Ошибка", "Сначала выберите путь для сохранения файла")
            return

        doc = Document()

        # Добавление заголовка "Список использованных источников"
        heading = doc.add_paragraph()
        run = heading.add_run("Список использованных источников")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.bold = True
        heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Добавление нумерованных источников
        for i, ref in enumerate(self.references, start=1):
            source = f"{i}. {ref}"
            para = doc.add_paragraph(source)
            para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            para.paragraph_format.left_indent = Cm(1.25)
            para.paragraph_format.line_spacing = 1.5
            run = para.runs[0]
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)

        # Настройка полей страницы
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(3)

        try:
            doc.save(self.save_path)
            messagebox.showinfo("Готово", "Источники успешно добавлены и документ сохранен")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить документ: {e}")

    def add_reference(self):
        source_type = self.source_type_var.get()
        author = self.author_var.get()
        title = self.title_var.get()
        if not source_type or not author or not title:
            messagebox.showerror("Ошибка", "Заполните все поля")
            return
        if source_type.lower() == "книга":
            reference = self.get_book_reference(author, title)
        else:
            reference = f"{source_type} - {author} - {title}"
        self.references.append(reference)
        messagebox.showinfo("Добавлено", "Источник добавлен в список")

        # Очистка полей ввода
        self.source_type_var.set("")
        self.author_var.set("")
        self.title_var.set("")

    def get_book_reference(self, author, title):
        api_key = "AIzaSyAHsGmRT8bOhXNPcA6Gb-A8QjzBXAaSd4Q"
        url = f"https://www.googleapis.com/books/v1/volumes?q=intitle:{title}+inauthor:{author}&key={api_key}"
        response = requests.get(url)
        if response.status_code != 200:
            return f"Книга - {author} - {title} (Не удалось получить информацию)"

        data = response.json()
        if "items" not in data or len(data["items"]) == 0:
            return f"Книга - {author} - {title} (Информация не найдена)"

        book = data["items"][0]["volumeInfo"]
        authors = book.get("authors", [])
        title = book.get("title", "")
        publisher = book.get("publisher", "")
        published_date = book.get("publishedDate", "")
        page_count = book.get("pageCount", "")

        # Извлечение года из даты публикации
        year = published_date.split("-")[0] if published_date else ""

        if len(authors) == 1:
            authors_str = authors[0]
        elif 1 < len(authors) <= 3:
            authors_str = ", ".join(authors)
        elif len(authors) > 3:
            authors_str = f"{authors[0]} [и др.]"
        else:
            authors_str = ""

        # Попытка получить город издания из API Google Books
        city = ""
        if "publishedCity" in book:
            city = book["publishedCity"]
        else:
            # Если город издания не указан, пытаемся найти информацию о книге в других источниках
            city = self.get_city_from_other_sources(author, title)

        # Если город издания не найден, используем город по умолчанию
        if not city:
            city = "Москва"  # Можно выбрать другой город по умолчанию

        reference = f"{authors_str}. {title}. {city}: {publisher}, {year}. {page_count} c."
        return reference

    def get_city_from_other_sources(self, author, title):
        # Попробуем найти информацию о книге в онлайн-библиотеке OpenLibrary
        url = f"https://openlibrary.org/search.json?author={author}&title={title}"
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            if "docs" in data and len(data["docs"]) > 0:
                doc = data["docs"][0]
                publish_places = doc.get("publish_places", [])
                if publish_places:
                    return publish_places[0]  # Возвращаем первое место издания, если оно доступно

        # Если информация не найдена в OpenLibrary, можно попробовать другие источники
        # Например, каталог WorldCat или другие онлайн-библиотеки

        # Если ничего не найдено, возвращаем None
        return None

if __name__ == "__main__":
    root = tk.Tk()
    app = ReferenceApp(root)

    # Кнопка для добавления источника в список
    add_button = tk.Button(root, text="Добавить источник", command=app.add_reference)
    add_button.pack()

    root.mainloop()