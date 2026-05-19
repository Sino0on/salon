# Использование официального образа Python в качестве базового
FROM python:3.13-slim

# Установка переменных окружения
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Установка рабочей директории
WORKDIR /app

# Установка зависимостей
COPY requirements.txt /app/
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

# Копирование проекта
COPY . /app/

# Сборка статических файлов
RUN python manage.py collectstatic --noinput

# Создание директории для отчетов
RUN mkdir -p scraper_reports


# Запуск gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8012", "--workers", "3", "--timeout", "120", "web_scraper.wsgi:application"]
