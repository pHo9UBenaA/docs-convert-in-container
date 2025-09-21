# ---- build stage ----
FROM mcr.microsoft.com/dotnet/sdk:8.0-bookworm-slim AS build
WORKDIR /src
COPY scripts/pptx-xml-to-jsonl/ ./

RUN dotnet publish pptx-xml-to-jsonl.csproj \
    -c Release \
    -r linux-x64 \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:PublishTrimmed=true \
    -o /out

# ---- runtime stage ----
FROM debian:bookworm-slim

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

COPY --from=build /out/ /opt/pptx-xml-to-jsonl/
ENV PATH="/opt/pptx-xml-to-jsonl:${PATH}"

WORKDIR /scripts
