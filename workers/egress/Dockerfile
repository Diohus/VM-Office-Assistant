# Dịch vụ egress (Q135, R17): proxy chỉ cho CONNECT api.openai.com:443. Không giải mã TLS, không giữ khóa. Chỉ thư viện chuẩn.
FROM python:3.11-slim@sha256:db20d8506e2f15b6db39b40b6015ca46bae855dec6dd326d462cfe79ae972893
COPY workers/egress/egress.py /opt/egress/egress.py
RUN useradd -u 1000 -m vm
USER 1000
WORKDIR /opt/egress
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
CMD ["python", "egress.py"]
