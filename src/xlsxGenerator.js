// ============ Excel 生成 (xlsx-populate) ============
// ひな形 public/template.xlsx を読み込んで、グリッドデータを流し込み、
// 罫線・結合・列幅などの書式を一切壊さずに新しい xlsx を生成する。
//
// xlsx-populate は「読み込んだXMLを最小限の差分だけ書き換える」設計のため
// ExcelJS で起きる "結合セル・印刷設定が壊れる" 問題を回避できる。

import XlsxPopulate from 'xlsx-populate/browser/xlsx-populate';

// === 設定 ===
const TEMPLATE_URL = `${import.meta.env.BASE_URL}template.xlsx`;
const DEFAULT_SHEET_NAME = 'ベース'; // 後で日付ごとのシートに切り替え予定

// === 色定義 ===
const COLOR_WORKING = '00FF00';   // 出勤時間ベース色 (緑)
// サービス種別ごとの利用者名セル色（出勤時間の緑を上書き）
// 後日 身体2 / 生活1 等を追加する場合はここに足す
const SERVICE_COLOR = {
  '身1': 'CC99FF',  // 身体1 → 薄紫
};

// === ヘルパー: 1-indexed の列番号から Excel の列文字 (A,B,...,AA,AB,...) に変換 ===
function colNumToLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// === ヘルパー: 時刻文字列 "7:00" → 15分スロットインデックス (0〜95) ===
function timeToSlot(timeStr) {
  if (!timeStr) return null;
  const m = String(timeStr).match(/(\d{1,2})[:時](\d{1,2})/);
  if (!m) return null;
  const h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  return h * 4 + Math.floor(min / 15);
}

// === ヘルパー: isStart のサービスが何スロット使うかを計算 ===
// 1枠 = 1〜15分以内なので、25分のサービスは 2枠を使う。
// staffGrid 上では「isStart=true の先頭セル + isStart=false の後続セル」の連続として表現される。
// 次の isStart に出会うか staffGrid に値が無くなったら終了。
function calcServiceLen(staffGrid, startIdx) {
  let len = 1;
  while (startIdx + len < 96) {
    const next = staffGrid[startIdx + len];
    if (!next || next.isStart) break;
    len++;
  }
  return len;
}

// === メイン: グリッドデータをひな形に流し込んで Excel をダウンロード ===
export async function generateXlsx({ workingStaff, staffShifts, staffTypes, grid, selectedDate }) {
  // 1. ひな形を fetch
  const res = await fetch(TEMPLATE_URL);
  if (!res.ok) throw new Error(`ひな形の読み込みに失敗: ${res.status}`);
  const buf = await res.arrayBuffer();

  // 2. xlsx-populate で読み込み
  const workbook = await XlsxPopulate.fromDataAsync(buf);

  // 3. 書き込み先シートを取得（とりあえず「ベース」固定）
  const sheet = workbook.sheet(DEFAULT_SHEET_NAME);
  if (!sheet) {
    throw new Error(`シート「${DEFAULT_SHEET_NAME}」が見つかりません`);
  }

  // 4. 曜日を計算
  let dateStr = selectedDate || '';
  if (/^\d{1,2}\/\d{1,2}$/.test(dateStr)) {
    const now = new Date();
    const [mo, da] = dateStr.split('/');
    dateStr = `${now.getFullYear()}-${String(mo).padStart(2, '0')}-${String(da).padStart(2, '0')}`;
  }
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  let weekday = '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const d = new Date(dateStr);
    if (!isNaN(d.getTime())) weekday = weekdays[d.getDay()];
  }

  // 5. スタッフごとに 上段(利用者名)+下段(区分+サービスコード) の2行を書き込み
  // 既存ヘッダ行(Row 1)は触らない
  let rowIdx = 2;
  workingStaff.forEach((staff, idx) => {
    const staffType = staffTypes[staff] || '社員';
    const staffGrid = grid[staff] || [];
    const shift = staffShifts[staff] || {};

    // ----- 上段: 利用者名行 -----
    if (idx === 0 && weekday) {
      sheet.cell(`A${rowIdx}`).value(weekday);
    }
    sheet.cell(`B${rowIdx}`).value(staff);

    // (1) 出勤時間の範囲を緑(#00FF00)で塗る (ベース色)
    const startSlot = timeToSlot(shift.start);
    const endSlot = timeToSlot(shift.end);
    if (startSlot != null && endSlot != null && endSlot > startSlot) {
      for (let i = startSlot; i < endSlot; i++) {
        const colLetter = colNumToLetter(3 + i);
        sheet.cell(`${colLetter}${rowIdx}`).style('fill', COLOR_WORKING);
      }
    }

    // (2) 利用者名 + サービス色 (全枠を塗る、利用者名は先頭枠のみ)
    for (let i = 0; i < 96; i++) {
      const cell = staffGrid[i];
      if (cell && cell.isStart) {
        const len = calcServiceLen(staffGrid, i);
        // 先頭枠に利用者名
        if (cell.user) {
          sheet.cell(`${colNumToLetter(3 + i)}${rowIdx}`).value(cell.user);
        }
        // サービス色を全枠に塗る (出勤時間の緑を上書き)
        const overrideColor = SERVICE_COLOR[cell.code];
        if (overrideColor) {
          for (let j = 0; j < len; j++) {
            sheet.cell(`${colNumToLetter(3 + i + j)}${rowIdx}`).style('fill', overrideColor);
          }
        }
      }
    }
    rowIdx++;

    // ----- 下段: 区分 + サービスコード行 (塗りなし、複数枠は結合) -----
    sheet.cell(`B${rowIdx}`).value(staffType);
    for (let i = 0; i < 96; i++) {
      const cell = staffGrid[i];
      if (cell && cell.isStart && cell.code) {
        const len = calcServiceLen(staffGrid, i);
        const startCol = colNumToLetter(3 + i);
        sheet.cell(`${startCol}${rowIdx}`).value(cell.code);
        if (len > 1) {
          const endCol = colNumToLetter(3 + i + len - 1);
          sheet.range(`${startCol}${rowIdx}:${endCol}${rowIdx}`).merged(true);
        }
      }
    }
    rowIdx++;
  });

  // 6. Blob としてエクスポート → ダウンロード
  const blob = await workbook.outputAsync('blob');
  const filename = buildFilename(dateStr);
  triggerDownload(blob, filename);
}

function buildFilename(dateStr) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    return `シフト表_${dateStr}.xlsx`;
  }
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `シフト表_${y}-${m}-${d}.xlsx`;
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
