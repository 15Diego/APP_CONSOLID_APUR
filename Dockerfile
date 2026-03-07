# Estágio 1: Build
FROM python:3.10-slim as builder

WORKDIR /app

# Instalar dependências de build
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Instalar dependências Python
COPY requirements.txt .
RUN pip install --user --no-cache-dir -r requirements.txt

# Estágio 2: Final
FROM python:3.10-slim

WORKDIR /app

# Instalar dependências de runtime (curl para healthcheck)
RUN apt-get update && apt-get install -y \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copiar dependências instaladas do estágio de build
COPY --from=builder /root/.local /root/.local
ENV PATH=/root/.local/bin:$PATH

# Copiar código da aplicação
COPY app.py .
COPY consolidate_relatorio_base.py .
COPY .streamlit .streamlit/

# Expor porta do Streamlit
EXPOSE 8501

# Healthcheck
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Comando para iniciar a aplicação
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
