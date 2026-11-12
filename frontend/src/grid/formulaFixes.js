// N9 (R-1015): sửa vài hàm của bộ tính Univer 1.3.1 cho giống Excel. Chỉ univerAdapter.js dùng file này.
// Đo bằng fixtures/templates/GIA-X20_thong_ke_vung_it_so.xlsx (Excel thật tính và lưu, tools/excel_check/make_stat_empty.py):
// - STDEV, STDEV.S, VAR, VAR.S, STDEVA, VARA: vùng không có số thì Univer ra 0, Excel ra #DIV/0!.
//   Biểu mẫu QC trống có IFERROR(STDEV(...),"") nên lưới hiện 0 thay vì trống, và PDF bị từ chối vì lệch 6 ô.
// - MINA, MAXA: Univer coi ô trống trong vùng là 0 (MINA của vùng 5.5 và ô trống ra 0); Excel bỏ qua ô trống.
// Nhóm .P (STDEV.P, VAR.P, STDEVP, VARP, STDEVPA, VARPA) và các hàm khác trong file đo đã khớp Excel, không đổi.
// Hàm đăng ký qua formula.function được đặt sau hàm có sẵn cùng tên, nên thay hàm đó (engine-formula: registerExecutors).
import { BaseFunction, ErrorType, ErrorValueObject, NumberValueObject } from '@univerjs/preset-sheets-core';

function spread({ sample, root, withText }) {
  return class extends BaseFunction {
    constructor(...args) {
      super(...args);
      this.minParams = 1;
      this.maxParams = 255;
    }

    calculate(...variants) {
      // Cách lấy số giữ nguyên như Univer (ô trống luôn bỏ; chữ, TRUE/FALSE trong vùng chỉ tính ở bản A).
      const flat = this.flattenArray(variants, !withText);
      if (flat.isError()) return flat;
      let n = 0;
      flat.iterator((v) => {
        if (v != null && !v.isNull()) n += 1;
      });
      // Mẫu cần ít nhất 2 số, toàn bộ cần ít nhất 1 số; thiếu thì Excel trả #DIV/0!.
      if (n < (sample ? 2 : 1)) return ErrorValueObject.create(ErrorType.DIV_BY_ZERO);
      const ddof = sample ? 1 : 0;
      return root ? flat.std(ddof) : flat.var(ddof);
    }
  };
}

function extreme(isMin) {
  return class extends BaseFunction {
    constructor(...args) {
      super(...args);
      this.minParams = 1;
      this.maxParams = 255;
    }

    calculate(...variants) {
      let best = null;
      const take = (x) => {
        if (best === null || (isMin ? x < best : x > best)) best = x;
      };
      for (const variant of variants) {
        if (variant.isNull()) continue;
        if (variant.isError()) return variant;
        if (variant.isArray()) {
          let error = null;
          variant.iterator((v) => {
            if (v == null || v.isNull()) return true; // ô trống: bỏ qua như Excel
            if (v.isError()) {
              error = v;
              return false;
            }
            if (v.isString()) take(0); // chữ trong vùng tính là 0 (bản A)
            else if (v.isBoolean()) take(v.getValue() ? 1 : 0);
            else take(Number(v.getValue()));
            return true;
          });
          if (error) return error;
          continue;
        }
        const one = variant.isString() || variant.isBoolean() ? variant.convertToNumberObjectValue() : variant;
        if (one.isError()) return one;
        take(Number(one.getValue()));
      }
      return NumberValueObject.create(best ?? 0);
    }
  };
}

const StdevS = spread({ sample: true, root: true, withText: false });
const VarS = spread({ sample: true, root: false, withText: false });

export const FORMULA_FIXES = [
  [StdevS, 'STDEV'],
  [StdevS, 'STDEV.S'],
  [VarS, 'VAR'],
  [VarS, 'VAR.S'],
  [spread({ sample: true, root: true, withText: true }), 'STDEVA'],
  [spread({ sample: true, root: false, withText: true }), 'VARA'],
  [extreme(true), 'MINA'],
  [extreme(false), 'MAXA'],
];
