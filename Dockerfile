FROM python:3.11-slim

WORKDIR /app

RUN pip install pandas psycopg2-binary boto3 openpyxl

COPY import_csv.py /app/import_csv.py

CMD ["python", "/app/import_csv.py"]
