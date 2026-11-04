// Giao diện lưới dùng cho phần nghiệp vụ (Q20). Đổi thư viện lưới thì chỉ thay file adapter, không sửa nơi gọi.
//
// createGrid(container, model) trả về:
//   compareWithSaved(): Promise<[{ sheet, cell, saved, computed }]>  ô công thức tính ra khác giá trị đã lưu
//   goTo(sheetName, a1)                                              chuyển sheet, chọn ô
//   visibleSheets(): string[]                                        tên sheet người dùng thấy
//   value(sheetName, a1)                                             giá trị thô đã tính
//   destroy()
export { createGrid } from './univerAdapter.js';
