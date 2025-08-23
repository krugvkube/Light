#!/bin/zsh

# Команда
APPLESCRIPT_CONTENT='do shell script "python3 " & quoted form of "___"'

# Создаем папку если нет
mkdir -p ~/Library/Script\ Libraries/

# Временный файл
TEMP_FILE=$(mktemp)
echo "$APPLESCRIPT_CONTENT" > "$TEMP_FILE.applescript"

# Компилируем
osacompile -o ~/Library/Script\ Libraries/PythonRunner.scpt "$TEMP_FILE.applescript"

# Удаляем временный файл
rm "$TEMP_FILE.applescript"

echo "AppleScript создан в ~/Library/Script Libraries/"
