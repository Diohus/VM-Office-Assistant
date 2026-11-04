# VM-Office-Assistant

Bản thử nghiệm của một công cụ hỗ trợ giấy tờ văn phòng: mở biểu mẫu Excel trong trình duyệt và giữ nguyên file gốc. Các chức năng khác đang được phát triển.

**Chỉ dùng dữ liệu giả.** Đây chưa phải sản phẩm hoàn chỉnh. Không tải lên tài liệu có thông tin người bệnh hay dữ liệu thật. Mọi file trong `fixtures/` đều là mẫu giả.

## Chạy thử

Cần Docker Desktop và Git Bash trên Windows.

```sh
sh tools/dev/up.sh
```

Mở http://127.0.0.1:8088. Tắt bằng `docker compose -p vm-office --env-file .local/compose.env down`.

App chỉ mở trên máy đang chạy. Mật khẩu cơ sở dữ liệu tự sinh vào `.local/`, không nằm trong kho.

## Kiểm thử

```sh
sh tests/integration/run.sh
```

Cần Python 3.11, Node 24 và Edge, cùng môi trường Python ở `workers/reader/.venv` có các gói trong `workers/reader/requirements.in` và `pytest`. Test chạy trên bộ container riêng ở cổng 8089.

## Báo lỗi

Tạo issue trong kho này, nêu các bước làm và điều xảy ra. Chỉ đính kèm file giả, không gửi file hay ảnh có dữ liệu thật. Nếu phát hiện lỗi bảo mật, đừng nêu chi tiết trong issue công khai, hãy báo trực tiếp cho người giữ kho.
