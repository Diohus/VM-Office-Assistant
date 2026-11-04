# Bước build giao diện, rồi phục vụ file tĩnh và chuyển /api cho API. Web là thành phần duy nhất mở cổng ra ngoài.
FROM node:24-slim@sha256:c86ecead9e55a5210ac1727de38a4ca840bf03dcaf34dc9749addbf383b1920c AS build
WORKDIR /src
COPY frontend/package.json frontend/package-lock.json ./
RUN --mount=type=secret,id=machine_ca sh -c 'NODE_EXTRA_CA_CERTS=/run/secrets/machine_ca npm ci --ignore-scripts --no-audit --no-fund'
COPY frontend/ ./
RUN npm run build

FROM nginxinc/nginx-unprivileged@sha256:0c79d56aee561a1d81c63f00eee5fb5fe29279560cdc55e91425133104c7fbe6
COPY frontend/nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /src/dist /usr/share/nginx/html
