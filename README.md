# Формирование отчетных реестров
Программа для автоматического формирования Реестров для отчета по форме Компании.
Формируются следующие реестры:
- формы плана/отчета за месяц/квартал для дальнейшего заполнения,
- готовые реестры согласно формы Компании в соответствии с приложенными отчетными файлами.

## Технологии
- Python
- pywin32

## Использование
Для использования скачайте [папку](/finished%20program) с файлами:
- исполняемы файл **_registers.exe_**
- шаблон реестра **_template.xlsx_** 
- **_Инструкция_** по пользованию программой 

Далее следуйте перечисленным шагам в Инструкции к программе.

## Разработка

### Установка зависимостей
Для установки зависимостей в созданном виртуальном окружении, поочередно выполните команды:
```
python -m venv env
```
```
source venv/Scripts/activate
```
```
python3 -m pip install --upgrade pip
```
```
pip install -r requirements.txt
```

### Подготовка к работе
Переместите в рабочую директорию файл шаблона реестра **_template.xlsx_** 

## Автор проекта
[Shamil Khamidullin](https://github.com/skhfh)
