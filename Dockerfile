FROM python:3.11-slim

WORKDIR /app

# Устанавливаем необходимые зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Устанавливаем необходимые системные пакеты
RUN apt-get update && apt-get install -y \
    libxml2-dev \
    libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

# Копируем файлы проекта
COPY word_server.py .
COPY pyproject.toml .
COPY LICENSE .
COPY README.md .

# Создаем MCP конфигурацию
RUN mkdir -p /root/.mcp
COPY mcp-config.json /root/.mcp/config.json

# Создаем директорию для документов
RUN mkdir -p /app/documents
VOLUME /app/documents

# Указываем переменную окружения для платформы
ENV PYTHONUNBUFFERED=1

# Запускаем сервер при старте контейнера
CMD ["python", "word_server.py"] 