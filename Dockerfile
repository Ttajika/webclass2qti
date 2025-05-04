FROM python:3.10-slim

# Pandoc &必要パッケージのインストール
RUN apt-get update && \
    apt-get install -y pandoc && \
    pip install --no-cache-dir streamlit pandas

# 作業ディレクトリ
WORKDIR /app

# アプリ本体をコピー
COPY . .

# ポート番号（Streamlitは8501を使用）
EXPOSE 8080

# 起動コマンド
CMD ["streamlit", "run", "conv.py", "--server.port=8080", "--server.enableCORS=false"]
