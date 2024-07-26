FROM python:3.11-slim

WORKDIR /app
COPY . /app
RUN pip install poetry flask
RUN poetry config virtualenvs.create false && poetry install --extras "http" --no-interaction --no-ansi

EXPOSE 8080
CMD ["python", "app.py"]
