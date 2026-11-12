# Dịch vụ llm (Q135): chỉ gọi model. Key chỉ vào qua Docker secret dạng file, không qua biến môi trường. Chỉ thư viện chuẩn.
FROM python:3.11-slim@sha256:db20d8506e2f15b6db39b40b6015ca46bae855dec6dd326d462cfe79ae972893
# CA của máy (máy có proxy công ty) gộp với CA hệ thống để xác minh TLS, không tắt xác minh (Q85). Đây là chứng chỉ công khai, không phải khóa.
RUN --mount=type=secret,id=machine_ca sh -c 'mkdir -p /opt/llm && cat /etc/ssl/certs/ca-certificates.crt /run/secrets/machine_ca > /opt/llm/ca.pem'
COPY workers/llm/llm_service.py /opt/llm/llm_service.py
RUN useradd -u 1000 -m vm
USER 1000
WORKDIR /opt/llm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
CMD ["python", "llm_service.py"]
