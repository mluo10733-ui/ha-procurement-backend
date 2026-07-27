FROM python:3.12-slim
WORKDIR /app
COPY inquiry/requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY inquiry/ .
ENV PYTHONUNBUFFERED=1
CMD ["python", "server.py"]
