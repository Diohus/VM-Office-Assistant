# API: không có mạng ra ngoài (compose), không giữ khóa AI, chạy không root, gốc chỉ đọc.
FROM python:3.11-slim@sha256:db20d8506e2f15b6db39b40b6015ca46bae855dec6dd326d462cfe79ae972893
COPY backend/requirements.lock /tmp/requirements.lock
# CA của máy chỉ dùng lúc build, không nằm trong image. Gói được kiểm theo hash trong lockfile.
RUN --mount=type=secret,id=machine_ca sh -c 'cat /etc/ssl/certs/ca-certificates.crt /run/secrets/machine_ca > /tmp/ca.pem && pip install --no-cache-dir --root-user-action=ignore --cert /tmp/ca.pem --require-hashes --no-deps -r /tmp/requirements.lock && rm /tmp/ca.pem /tmp/requirements.lock'
COPY backend/app /opt/app/app
RUN useradd -u 1000 -m vm && mkdir -p /store /spool /lospool && chown 1000:1000 /store /spool /lospool
USER 1000
WORKDIR /opt/app
ENV PYTHONDONTWRITEBYTECODE=1
CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--no-server-header", "--no-access-log"]
