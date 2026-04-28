FROM python:3.11-slim-bullseye

WORKDIR /app

RUN apt-get update && apt-get install -y \
    libaio1 \
    unzip \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# copiar o instant client
COPY instantclient*.zip /opt/

RUN cd /opt && \
    unzip instantclient*.zip && \
    rm instantclient*.zip

# detectar automaticamente pasta (melhor prática)
ENV LD_LIBRARY_PATH=/opt/instantclient_23_26

COPY app/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app/ .

EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
