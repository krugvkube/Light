import os

file_path_Copy = str(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))+"/VBA"+"/CopyToEveryList.txt" 
file_path_Copy2 = str(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))+"/VBA"+"/SaveTest.txt"
file_path_Light = str(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

def process_file():
    try:
        # Открываем файл для чтения
        with open(file_path_Copy, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        
        # Обрабатываем строки
        modified_lines = []
        for line in lines:
            if 'filePath = ' in line:
                # Заменяем только в строках с filePath =
                modified_line = line.replace('&&&&', file_path_Light)
                modified_lines.append(modified_line)
            else:
                # Оставляем другие строки без изменений
                modified_lines.append(line)
        
        # Записываем изменения обратно в файл
        with open(file_path_Copy, 'w', encoding='utf-8') as file:
            file.writelines(modified_lines)
        
        print("Файл успешно обработан!")
        
    except FileNotFoundError:
        print("Ошибка: файл CopyToEveryList.txt не найден")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

# Запускаем обработку
if __name__ == "__main__":
    process_file()

def process_file2():
    try:
        # Открываем файл для чтения
        with open(file_path_Copy2, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        
        # Обрабатываем строки
        modified_lines = []
        for line in lines:
            if 'filePath = ' in line:
                # Заменяем только в строках с filePath =
                modified_line = line.replace('&&&&', file_path_Light)
                modified_lines.append(modified_line)
            else:
                # Оставляем другие строки без изменений
                modified_lines.append(line)
        
        # Записываем изменения обратно в файл
        with open(file_path_Copy2, 'w', encoding='utf-8') as file:
            file.writelines(modified_lines)
        
        print("Файл успешно обработан!")
        
    except FileNotFoundError:
        print("Ошибка: файл SaveTest.txt не найден")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

# Запускаем обработку
if __name__ == "__main__":
    process_file2()
