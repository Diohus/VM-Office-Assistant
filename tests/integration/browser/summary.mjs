// In tóm tắt ../out/browser.json để xem nhanh khi chạy tay.
import fs from 'node:fs';

const r = JSON.parse(fs.readFileSync('../out/browser.json', 'utf8'));
console.log('outbound', r.outbound, 'errs', r.page_errors, 'csp', r.csp_violations);
for (const [k, v] of Object.entries(r.files)) {
  console.log(k, v.seconds, 's f=', v.formulas, 'sheets', v.sheets.join(','), 'lech', v.mismatches.length, JSON.stringify(v.mismatches.slice(0, 5)), 'prob', JSON.stringify(v.problems).slice(0, 300));
  console.log('   ', v.notices.map((n) => n.text).join(' | '));
}
const { files, ...rest } = r;
console.log(JSON.stringify(rest).slice(0, 2500));
const q = r.files['GIA-QC-01_mau.xlsx'];
console.log(q.comment_text, q.jump_comment, q.jump_chart, q.after_close, q.after_reopen);
const d = r.files['GIA-QC-01_da_dien.xlsx'];
console.log(d.B39, d.cf_rules, d.dv_H8);
