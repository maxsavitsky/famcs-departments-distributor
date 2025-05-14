## Описание
Данный скрипт распределяет студентов по их среднему баллу и приоритетам между кафедрами (в данной реализации на специальности "Информатика")  
Для внесения своего списка измените структуру `departments`  
[Пример таблицы](https://docs.google.com/spreadsheets/d/1_Qq95tMhy7mlyoYDWhpIQ7HgNZo0QeIPELi9ISYQecI/edit?usp=sharing)

## Авторизация
Для работы скрипта необходимо создать сервисный аккаунт Google Cloud, подключить туда Google Sheets API и загрузить файл `credentials.json`  
Плюс-минус хорошая инструкция описана [здесь](https://habr.com/ru/articles/575160/)
Так же нужно указать ссылку на вашу таблицу вот этой строчке:
```python
spreadsheet = gspread.service_account(filename='credentials.json').open_by_url("URL")
```