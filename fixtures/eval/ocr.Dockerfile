# Image chỉ để đo: Tesseract chạy không mạng (Q83). Image sản phẩm là việc của Dev (T-DEV-02).
# Dựng: docker build -t vmai-ocr-do -f fixtures/eval/ocr.Dockerfile fixtures/eval
FROM debian:bookworm-slim
RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-vie tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*
USER nobody
WORKDIR /work
