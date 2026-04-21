# Node.js + Python ハイブリッド構成
FROM node:20-slim

# Python 3 と必要なビルドツールをインストール
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    python3-venv \
    && rm -rf /var/lib/apt/lists/*

# pip 用のシンボリックリンク
RUN ln -sf /usr/bin/python3 /usr/local/bin/python3 \
    && ln -sf /usr/bin/pip3 /usr/local/bin/pip3

WORKDIR /app

# Node.js 依存
COPY package*.json ./
RUN npm ci --omit=dev

# Python 依存（venv 経由）
COPY requirements.txt ./
RUN python3 -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"
RUN pip install --no-cache-dir -r requirements.txt

# アプリコードをコピー
COPY . .

# output ディレクトリを作成
RUN mkdir -p /app/output

# 環境変数
ENV NODE_ENV=production
ENV PORT=10000

EXPOSE 10000

CMD ["node", "server.js"]
