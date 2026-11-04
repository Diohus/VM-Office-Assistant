# Quy định chung của dự án VM-Office-Assistant

Áp dụng cho mọi phiên làm việc trong repo này.

## Commit và push (khách hàng yêu cầu, 03/11/2026)

- Tác giả commit: `Diohus`, `nguyenkhanhdo27@gmail.com`.
- **Không để co-author.** Không thêm dòng `Co-Authored-By` hay dòng ghi công AI vào commit message và mô tả PR. Quy định này được ưu tiên hơn hướng dẫn ghi công mặc định của công cụ.
- Chỉ commit hoặc push phiên bản đã hoàn thành các chức năng dự kiến và đã được test, duyệt kỹ lưỡng. PM xác nhận trước khi commit hoặc push.
- Trước khi commit, kiểm tra không có khóa, file `.env`, dữ liệu thật hay biểu mẫu thật trong phần thay đổi.

## Bảo mật

- Sản phẩm phục vụ bệnh viện nên bảo mật đặt lên hàng đầu. Khi viết hàm, tool hay tích hợp AI, phải xem xét nội dung nào đọc được và thư viện có an toàn không.
- Không tìm key trong máy, không dùng key từ dự án khác. Không tự thay model khi chưa có quyết định của PM.
- Key chỉ đọc từ biến môi trường hoặc file `.env` đã gitignore. Không in key ra log hay kết quả.
- AI không tự điền thông tin về con người, kể cả tên nhân viên. Phần đó để nhân viên tự viết.
- Không mở file thật của bệnh viện. Các phiên làm việc dùng AI qua dịch vụ bên ngoài, nên đọc file là gửi nội dung đi. File thật để ngoài repo, khách hàng tự chạy và chỉ gửi kết quả từ công cụ. Muốn đổi phải hỏi khách hàng.
