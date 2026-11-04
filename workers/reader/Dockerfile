# Bộ đọc file tải lên: không mạng, không khóa, không root, gốc chỉ đọc (compose). Chỉ trao đổi qua thư mục spool.
FROM python:3.11-slim@sha256:db20d8506e2f15b6db39b40b6015ca46bae855dec6dd326d462cfe79ae972893
COPY workers/reader/requirements.lock /tmp/requirements.lock
RUN --mount=type=secret,id=machine_ca sh -c 'cat /etc/ssl/certs/ca-certificates.crt /run/secrets/machine_ca > /tmp/ca.pem && pip install --no-cache-dir --root-user-action=ignore --cert /tmp/ca.pem --require-hashes --no-deps -r /tmp/requirements.lock && rm /tmp/ca.pem /tmp/requirements.lock'
COPY workers/reader/reader /opt/app/reader
COPY vmai/src/vmai /opt/app/vmai
RUN useradd -u 1000 -m vm && mkdir -p /spool && chown 1000:1000 /spool
USER 1000
WORKDIR /opt/app
ENV PYTHONPATH=/opt/app PYTHONDONTWRITEBYTECODE=1
CMD ["python", "-m", "reader.worker"]
