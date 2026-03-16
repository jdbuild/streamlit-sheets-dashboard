FROM python:3.12-slim

WORKDIR /app

COPY pyproject.toml /app/
COPY src /app/src
COPY app.py /app/
COPY scripts /app/scripts
COPY data /app/data

RUN pip install --no-cache-dir .

ENV APP_STATE_DIR=/app/app_state
ENV PYTHONUNBUFFERED=1
EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
