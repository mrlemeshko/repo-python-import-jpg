import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from io import BytesIO

# Функция для поиска изображения по запросу
def get_image_from_query(query):
    # Замените этот URL на URL вашего API поиска изображений или на сайт, с которого нужно парсить изображения
    search_url = f"https://source.unsplash.com/100x100/?{query}"
    response = requests.get(search_url)
    
    if response.status_code == 200:
        return Image(BytesIO(response.content))
    else:
        return None

# Открываем существующий файл Excel
file_path = '/mnt/data/Тестовые позиции для парсинга.xlsx'
wb = load_workbook(filename=file_path)
ws = wb.active

# Проходим по строкам, где находятся наименования товаров
for row in range(8, 16):  # D8:D15
    item_name = ws[f'D{row}'].value
    if item_name:
        print(f"Ищем изображение для: {item_name}")
        img = get_image_from_query(item_name)
        if img:
            ws.add_image(img, f'E{row}')  # Вставляем изображение в колонку E напротив наименования
        else:
            print(f"Не удалось найти изображение для: {item_name}")

# Сохраняем обновленный файл
wb.save('/mnt/data/Тестовые позиции для парсинга_with_images.xlsx')
