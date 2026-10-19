# Financial Dashboard

Веб-приложение на Flask для анализа российской бухгалтерской отчётности в Excel: баланс, финансовые результаты, показатели компаний и сводная таблица.

## Локальный запуск

Требуется Python 3.11 или новее.

```sh
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Открыть http://localhost:5000. Отчётность загружается через интерфейс; реальные Excel-файлы и база SQLite в Git не входят.

## Хранение и настройки

- `DB_PATH` — путь к SQLite, по умолчанию `data.db` рядом с приложением.
- `UPLOAD_DIR` — папка загрузок, по умолчанию `uploads/`.
- `REPORTS_DIR` — необязательная папка исходных Excel для первичного наполнения пустой базы.
- `PORT` — порт локального запуска, по умолчанию `5000`.
- `FLASK_DEBUG=1` включает локальную отладку.

`app.py` содержит маршруты и работу с SQLite, `parser.py` — разбор отчётности и расчёты, `templates/` и `static/` — интерфейс.

Docker использует Python 3.11 и Gunicorn на порту 8080. В `render.yaml` диск подключён к `/app/data`; для хранения SQLite на этом диске нужно явно задать `DB_PATH=/app/data/data.db`.
