import os
from reportlab.pdfgen import canvas
import random
import shutil

# Список имён и фамилий
names = [
    "Иванов", "Петров", "Сидоров", "Смирнов", "Кузнецов",
    "Попов", "Васильев", "Соколов", "Михайлов", "Новиков",
    "Фёдоров", "Морозов", "Волков", "Алексеев", "Лебедев",
    "Семенов", "Егоров", "Павлов", "Крылов", "Беляев",
    "Тарасов", "Борисов", "Зайцев", "Ильин", "Макаров",
    "Николаев", "Захаров", "Белов", "Медведев", "Антонов"
]
prev = ["AG", "AV", "RE"]
razdel = [" - ", " ", "."]
# Функция для создания PDF файла
def create_pdf(file_name, content, base_folder):
    file_path = os.path.join(base_folder, f"{file_name}.pdf")
    c = canvas.Canvas(file_path)
    c.drawString(100, 750, content)
    c.save()


def create_folder(base_folder, name):
    # Создаем папку для каждого имени/фамилии
    folder_path = os.path.join(base_folder, name)
    os.makedirs(folder_path, exist_ok=True)



def main():
    # Папка, куда будут сохраняться все созданные файлы
    base_folder = "Generated_FOLDER"
    path2 = "Generated_PDF"
    # Создаем базовую папку, если она не существует
    if not os.path.exists(base_folder):
        os.makedirs(base_folder)
    if not os.path.exists(path2):
        os.makedirs(path2)

    for name in names:
        file_name = random.choice(prev) + " " + random.choice(names) + random.choice(razdel) + random.choice(names)

        create_folder(base_folder, file_name)

        create_pdf(file_name, file_name, path2)


if __name__ == '__main__':
    # Запуск функции
    main()

    print("Папки и PDF файлы успешно созданы.")
