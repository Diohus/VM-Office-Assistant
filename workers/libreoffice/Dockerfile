# Worker LibreOffice: chạy với network none, gốc chỉ đọc, bỏ mọi capability, không root (compose). File tải lên có thể độc.
FROM debian:trixie-slim@sha256:f6b08b09a05725f9431502d3b5714f5c4576e82c1a84fa1debe0c24a7e0c3ec4
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-calc python3 fonts-liberation fonts-dejavu-core && rm -rf /var/lib/apt/lists/*
RUN useradd -u 1000 -m lo && mkdir -p /spool && chown 1000:1000 /spool
COPY workers/libreoffice/worker.py /opt/worker.py
USER 1000
WORKDIR /home/lo
CMD ["python3", "/opt/worker.py"]
