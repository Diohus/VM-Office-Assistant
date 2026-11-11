// Giao diện lưới dùng cho phần nghiệp vụ (Q20). Đổi thư viện lưới thì chỉ thay file adapter, không sửa nơi gọi.
//
// createGrid(container, model) trả về:
//   compareWithSaved(): Promise<[{ sheet, cell, saved, computed }]>  ô công thức tính ra khác giá trị đã lưu
//   goTo(sheetName, a1)                                              chuyển sheet, chọn ô
//   visibleSheets(): string[]                                        tên sheet người dùng thấy
//   value(sheetName, a1)                                             giá trị thô đã tính
//   settled(): Promise<{ ok, timedOut? }>                            chờ bộ tính xong; ok false là quá hạn, ô công thức có thể còn số cũ
//   isCalculating()                                                  bộ tính còn chạy (tùy chọn onCalculating báo khi đổi)
//   destroy()
export { createGrid } from './univerAdapter.js';
