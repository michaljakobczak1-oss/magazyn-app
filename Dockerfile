FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN mkdir -p data static/uploads
EXPOSE 5000
CMD gunicorn -b 0.0.0.0:${PORT:-5000} -w 1 app:app
