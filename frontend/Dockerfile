FROM python:3.11-slim

WORKDIR /app

RUN pip install --no-cache-dir streamlit==1.38.0 requests==2.32.3

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=8501"]
