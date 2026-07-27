FROM python:3.12-slim
WORKDIR /app
COPY reconciliation/requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY reconciliation/ .
ENV PYTHONUNBUFFERED=1
CMD ["python", "reconcile_app_v2.py"]
