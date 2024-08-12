import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from io import BytesIO
from PIL import Image as PILImage

# Функция для поиска изображения по запросу в Yandex
def get_image_from_yandex(query):
    # Используем запрос к Яндекс Картинкам
    search_url = f"https://yandex.com/images/search?text={query}&isize=gt&img_url=&iorient=square"
    
    # Пытаемся получить данные
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"}
    
    try:
        response = requests.get(search_url, headers=headers)
        response.raise_for_status()

        # Вытягиваем URL первой картинки из поисковой выдачи
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(response.text, "html.parser")
        img_tags = soup.find_all("img", {"class": "serp-item__thumb justifier__thumb"})
        
        if img_tags:
            img_url = img_tags[0]["src"]
            img_response = requests.get(img_url)
            img = PILImage.open(BytesIO(img_response.content))
            
            # Проверяем размер изображения
            if img.size[0] >= 100 and img.size[1] >= 100:
                return Image(BytesIO(img_response.content))
        return None
    except Exception as e:
        print(f"Ошибка при получении изображения для {query}: {e}")
        return None

# Открываем существующий файл Excel
file_path = 'Тестовые_позиции_для_парсинга.xlsx'
wb = load_workbook(filename=file_path)
ws = wb.active

# Проходим по строкам, где находятся наименования товаров (D1:D15)
for row in range(1, 16):  # Строки 1-15, в колонке D
    item_name = ws[f'D{row}'].value
    if item_name:
        print(f"Ищем изображение для: {item_name}")
        img = get_image_from_yandex(item_name)
        if img:
            ws.add_image(img, f'E{row}')  # Вставляем изображение в колонку E напротив наименования
        else:
            print(f"Не удалось найти изображение для: {item_name}")

# Сохраняем обновленный файл
output_file_path = 'Тестовые_позиции_для_парсинга_with_images.xlsx'
wb.save(output_file_path)

print(f"Изображения успешно добавлены в файл и сохранены как {output_file_path}!")
