import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from io import BytesIO
from PIL import Image as PILImage
from bs4 import BeautifulSoup  # Добавляем этот импорт

# Функция для поиска изображения по текстовому запросу через Яндекс
def get_image_from_yandex(query):
    search_url = 'https://yandex.ru/images/search'
    params = {'text': query}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"}
    
    try:
        response = requests.get(search_url, params=params, headers=headers)
        response.raise_for_status()

        # Парсинг HTML ответа и поиск изображения
        soup = BeautifulSoup(response.text, 'html.parser')
        img_tags = soup.find_all("img", {"class": "serp-item__thumb justifier__thumb"})

        if img_tags:
            img_url = img_tags[0]["src"]
            img_response = requests.get(img_url)
            img = PILImage.open(BytesIO(img_response.content))
            
            if img.size[0] >= 100 and img.size[1] >= 100:
                return Image(BytesIO(img_response.content))
        return None
    except Exception as e:
        print(f"Ошибка при получении изображения для {query}: {e}")
        return None

# Читаем данные из текстового файла
input_file_path = 'Тестовые_позиции.txt'
output_file_path = 'Тестовые_позиции_с_изображениями.xlsx'

try:
    with open(input_file_path, 'r', encoding='utf-8') as file:
        content = file.read()
        print(f"Файл успешно прочитан.")
        items = [item.strip() for item in content.split(',') if item.strip()]
        print(f"Найдено {len(items)} артикулов.")
except Exception as e:
    print(f"Ошибка при чтении файла: {e}")
    items = []

# Создаем новый Excel файл
wb = Workbook()
ws = wb.active
ws.title = "Результаты"

# Добавляем заголовки
ws['A1'] = 'Артикул'
ws['B1'] = 'Изображение'

# Проходим по каждому артикулу
for index, item_name in enumerate(items, start=2):
    if item_name:
        print(f"Ищем изображение для: {item_name}")
        ws.cell(row=index, column=1, value=item_name)  # Записываем артикул в колонку A
        img = get_image_from_yandex(item_name)
        if img:
            ws.add_image(img, f'B{index}')  # Вставляем изображение в колонку B напротив артикула
        else:
            print(f"Не удалось найти изображение для: {item_name}")
    else:
        print(f"Пропускаем пустую строку: {index-1}")

# Сохраняем файл Excel
wb.save(output_file_path)

print(f"Результаты сохранены в файл {output_file_path}!")
