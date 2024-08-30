import os

# Проверка, есть ли такой файл по указанному пути
def check_file_exists(file_path):
    return os.path.exists(file_path)