FROM node:20-slim

# Puppeteer (Chromium) uchun kerak bo'ladigan tizim kutubxonalari
RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates fonts-liberation fonts-noto-color-emoji \
    libasound2 libatk-bridge2.0-0 libatk1.0-0 libc6 libcairo2 libcups2 \
    libdbus-1-3 libexpat1 libfontconfig1 libgbm1 libglib2.0-0 libgtk-3-0 \
    libnspr4 libnss3 libpango-1.0-0 libx11-6 libxcomposite1 libxdamage1 \
    libxext6 libxfixes3 libxrandr2 libxkbcommon0 wget \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json ./
RUN npm install --no-audit --no-fund
COPY . .

CMD ["node", "bot.mjs"]
