#!/bin/sh
# Sinh lockfile Python có hash cho Linux, chạy trong đúng image mà sản phẩm dùng. Chạy từ gốc repo: sh tools/deps/lock.sh
set -eu
IMAGE=python:3.11-slim@sha256:db20d8506e2f15b6db39b40b6015ca46bae855dec6dd326d462cfe79ae972893
MSYS_NO_PATHCONV=1 docker run --rm -v "$(pwd -W 2>/dev/null || pwd):/w" -w /w "$IMAGE" sh -c '
  set -eu
  cat /etc/ssl/certs/ca-certificates.crt .local/machine-ca.pem > /tmp/ca.pem
  export PIP_CERT=/tmp/ca.pem
  pip install -q --root-user-action=ignore pip-tools==7.5.3
  pip-compile -q --generate-hashes --allow-unsafe --strip-extras --no-header --output-file=backend/requirements.lock backend/requirements.in
  pip-compile -q --generate-hashes --allow-unsafe --strip-extras --no-header --output-file=workers/reader/requirements.lock workers/reader/requirements.in
'
