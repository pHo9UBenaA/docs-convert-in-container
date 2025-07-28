FROM debian:bullseye-slim

# Install LibreOffice first (this layer will be cached)
RUN apt-get update && apt-get install -y \
    libreoffice \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install other required packages
RUN apt-get update && apt-get install -y \
    poppler-utils \
    fonts-noto-cjk \
    locales \
    gnumeric \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set Japanese locale
RUN sed -i '/ja_JP.UTF-8/s/^# //g' /etc/locale.gen && \
    locale-gen ja_JP.UTF-8
ENV LANG=ja_JP.UTF-8
ENV LC_ALL=ja_JP.UTF-8

WORKDIR /app
