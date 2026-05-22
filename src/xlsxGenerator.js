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

    // 上段: A=曜日(最初のスタッフのみ), B=氏名, C～CT(96スロット)=利用者名(isStartのみ)
    if (idx === 0 && weekday) {
      sheet.cell(`A${rowIdx}`).value(weekday);
    }
    sheet.cell(`B${rowIdx}`).value(staff);
    for (let i = 0; i < 96; i++) {
      const cell = staffGrid[i];
      if (cell && cell.isStart && cell.user) {
        const colLetter = colNumToLetter(3 + i); // C=3 始まり
        sheet.cell(`${colLetter}${rowIdx}`).value(cell.user);
      }
    }
    rowIdx++;

    // 下段: B=区分, C～CT=サービスコード(isStartのみ)
    sheet.cell(`B${rowIdx}`).value(staffType);
    for (let i = 0; i < 96; i++) {
      const cell = staffGrid[i];
      if (cell && cell.isStart && cell.code) {
        const colLetter = colNumToLetter(3 + i);
        sheet.cell(`${colLetter}${rowIdx}`).value(cell.code);
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
