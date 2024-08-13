from icrawler.builtin import GoogleImageCrawler
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
import os
from io import BytesIO
from PIL import Image as PILImage

# Папка для сохранения изображений
output_folder = 'downloaded_images'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Функция для загрузки изображений с Google и вставки их в Excel
def download_images_and_insert_to_excel(queries, max_num=1):
    # Создаем новый Excel файл
    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты"

    # Добавляем заголовки
    ws['A1'] = 'Артикул'
    ws['B1'] = 'Изображение'

    # Создаем экземпляр GoogleImageCrawler
    google_crawler = GoogleImageCrawler(storage={'root_dir': output_folder})

    for index, query in enumerate(queries, start=2):
        print(f"Ищем изображения для: {query}")
        # Запуск поиска изображений
        google_crawler.crawl(keyword=query, max_num=max_num)

        # Получаем первый загруженный файл
        images = os.listdir(output_folder)
        if images:
            img_path = os.path.join(output_folder, images[0])
            
            # Открываем изображение и вставляем его в Excel
            img = PILImage.open(img_path)
            buffer = BytesIO()
            img.save(buffer, format="JPEG")
            img_excel = ExcelImage(BytesIO(buffer.getvalue()))

            ws.cell(row=index, column=1, value=query)
            img_excel.anchor = f'B{index}'
            ws.add_image(img_excel)
            
            # Удаляем изображение после вставки
            os.remove(img_path)
        else:
            print(f"Не удалось найти изображения для: {query}")

    # Сохраняем файл Excel
    wb.save('Тестовые_позиции_с_изображениями.xlsx')
    print("Результаты сохранены в файл 'Тестовые_позиции_с_изображениями.xlsx'")

# Пример запросов
queries = ['iPhone 13', 'Samsung Galaxy S21', 'Sony PlayStation 5']

# Загрузка изображений и вставка в Excel
download_images_and_insert_to_excel(queries)
