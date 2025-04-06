FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY word_server.py .
COPY pyproject.toml .
COPY LICENSE .
COPY README.md .

# Create MCP configuration
RUN mkdir -p /root/.mcp
COPY mcp-config.json /root/.mcp/config.json

# Create directory for documents
RUN mkdir -p /app/documents

# Expose ports (if needed in the future)
# EXPOSE 8000

# Run server on container start
CMD ["python", "word_server.py"] 