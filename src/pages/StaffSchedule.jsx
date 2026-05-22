import React, { useState, useMemo, useRef, useEffect } from 'react';
import { generateXlsx } from '../xlsxGenerator';
import { loadStaff, lookupType, addStaff } from '../store/staffStore';

// ============ CSV パーサ ============
function parseCSV(text) {
  text = text.replace(/^\uFEFF/, '');
  const rows = [];
  let cur = '', row = [], inQuote = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuote) {
      if (c === '"' && text[i+1] === '"') { cur += '"'; i++; }
      else if (c === '"') { inQuote = false; }
      else { cur += c; }
    } else {
      if (c === '"') inQuote = true;
      else if (c === ',') { row.push(cur); cur = ''; }
      else if (c === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
      else if (c === '\r') {}
      else { cur += c; }
    }
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows.filter(r => r.some(v => v && v.trim() !== ''));
}

function detectColumns(header) {
  const map = { staff: -1, user: -1, date: -1, startTime: -1, endTime: -1, service: -1 };
  const patterns = {
    staff:     /(担当者|職員|スタッフ|ヘルパー|提供者|サービス提供者)/,
    user:      /(利用者|お客様|被介護者|患者)/,
    date:      /(日付|提供日|サービス提供日|実施日|年月日)/,
    startTime: /(開始)/,
    endTime:   /(終了|了)/,
    service:   /(サービス種別|サービス内容|種別|内容|区分)/
  };
  header.forEach((h, i) => {
    const v = (h || '').trim();
    if (!v) return;
    for (const key of Object.keys(patterns)) {
      if (map[key] === -1 && patterns[key].test(v)) { map[key] = i; return; }
    }
  });
  return map;
}

function parseTime(str) {
  if (!str) return null;
  str = String(str).trim();
  const m = str.match(/(\d{1,2})[:時](\d{1,2})/);
  if (!m) return null;
  const h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  if (h < 0 || h > 24 || min < 0 || min > 59) return null;
  return h * 60 + min;
}

// ============ PDF.js 動的ロード ============
let pdfJsPromise = null;
function loadPdfJs() {
  if (pdfJsPromise) return pdfJsPromise;
  pdfJsPromise = new Promise((resolve, reject) => {
    if (window['pdfjsLib']) { resolve(window['pdfjsLib']); return; }
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    s.onload = () => {
      const lib = window['pdfjsLib'];
      lib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(lib);
    };
    s.onerror = () => reject(new Error('PDF.js の読み込みに失敗しました'));
    document.head.appendChild(s);
  });
  return pdfJsPromise;
}

// ============ サービスコード変換 ============
// サービス名 → 短縮コード(身１/身２/生１/生２/通１/重訪Ⅱ/移動/家事/個別)
// この関数は上の方で定義済みのclassifyServiceとは別目的: Excel/HTML出力用の短縮コード
// 注: classifyService と toServiceCode は別関数として既存

// ============ TSV 生成(タブ区切り) ============
// テンプレ原本のA1セルに貼り付けるためのTSVを生成
// Excel/Googleスプレッドシートで Ctrl+V するとセルが自動展開される
// 注: TSVは結合や色を表現できないため、原本に書式が既にあれば書式を保持して値だけ流し込む
function generateTSV({ workingStaff, staffShifts, staffTypes, grid, selectedDate }) {
  // 曜日計算
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

  const lines = [];

  // === 1行目: ヘッダ ===
  // A1=曜日, B1=氏名, C1-F1=00:00:00, G1-J1=01:00:00 ...
  // TSVでは結合表現できないので、各時間ラベルは「1時間目の最初のセルだけ」に時刻、他は空
  const headerRow = ['曜日', '氏名'];
  for (let h = 0; h < 24; h++) {
    const label = `${String(h).padStart(2, '0')}:00:00`;
    headerRow.push(label);  // 4スロット中の最初に時刻
    headerRow.push('');     // 2スロット目
    headerRow.push('');     // 3スロット目
    headerRow.push('');     // 4スロット目
  }
  lines.push(headerRow.join('\t'));

  // === スタッフ行 ===
  workingStaff.forEach((staff, idx) => {
    const staffType = staffTypes[staff] || '社員';
    const staffGrid = grid[staff] || [];

    // 上段: 利用者名行(A列=曜日は最初のスタッフの上段だけ、それ以外は空)
    const upperRow = [idx === 0 ? weekday : '', staff];
    for (let i = 0; i < 96; i++) {
      const cell = staffGrid[i];
      if (cell && cell.isStart) {
        upperRow.push(cell.user);
      } else {
        upperRow.push('');
      }
    }
    lines.push(upperRow.join('\t'));

    // 下段: 区分 + サービスコード(コードは結合の開始セルだけ書く)
    const lowerRow = ['', staffType];
    for (let i = 0; i < 96; i++) {
      const cell = staffGrid[i];
      if (cell && cell.isStart && cell.code) {
        lowerRow.push(cell.code);
      } else {
        lowerRow.push('');
      }
    }
    lines.push(lowerRow.join('\t'));
  });

  return lines.join('\n');
}

// ============ クリップボードにテキスト(TSV)としてコピー ============
async function copyTextToClipboard(text) {
  // モダンブラウザのClipboard API
  if (navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      console.warn('clipboard.writeText failed, falling back:', e);
    }
  }
  // フォールバック: execCommand
  const temp = document.createElement('textarea');
  temp.value = text;
  temp.style.position = 'fixed';
  temp.style.left = '-9999px';
  document.body.appendChild(temp);
  temp.select();
  try {
    const ok = document.execCommand('copy');
    document.body.removeChild(temp);
    return ok;
  } catch (e) {
    document.body.removeChild(temp);
    return false;
  }
}



// ============ PDF → 行ベース テキスト復元 ============
// 同じY座標(±許容)のテキストを1行にまとめ、X座標順に並べる
async function pdfToLines(file) {
  const pdfjsLib = await loadPdfJs();
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const allLines = []; // { page, y, items: [{x, str}] }

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    // items: { str, transform: [a,b,c,d,e,f] } e=x, f=y
    const lineMap = new Map(); // y(丸め) -> items
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const x = it.transform[4];
      const y = Math.round(it.transform[5]); // 同一行検出のため整数化
      // 近傍Yをマージ(±2px許容)
      let bucket = null;
      for (const key of lineMap.keys()) {
        if (Math.abs(key - y) <= 2) { bucket = key; break; }
      }
      const k = bucket !== null ? bucket : y;
      if (!lineMap.has(k)) lineMap.set(k, []);
      lineMap.get(k).push({ x, str: it.str });
    }
    // Y降順(上から下) → X昇順
    const sortedY = [...lineMap.keys()].sort((a, b) => b - a);
    for (const y of sortedY) {
      const items = lineMap.get(y).sort((a, b) => a.x - b.x);
      allLines.push({ page: p, y, items, text: items.map(i => i.str).join(' ') });
    }
  }
  return allLines;
}

// ============ カイポケ「ヘルパー別シフト表」PDF専用抽出 ============
// 想定フォーマット:
//   ヘッダ行: "5月1日(金) 07:00 ～ 07:25 25 大川 富子様 身体介護１・夜"
//   スタッフ名は "ヘルパー名:" の次行 or 近傍に単独で出現する
function extractFromPdfLines(lines, filename) {
  const out = [];
  const debug = { lines: [], notes: [] };

  // ---------- スタッフ名検出 ----------
  // 「ヘルパー名:」を含む行を見つけ、その近辺で「姓 名」形式(漢字+空白+漢字)の行を探す
  // 除外: アムールケアステーション/事業所/対象期間/発行日/シフト表/予定/実績 等
  const excludeStaff = /(アムール|ケアステーション|事業所|対象期間|発行日|シフト表|予定|実績|シフト|月分|出力|印刷|備考|サービス|身体|生活|介護|提供時間|ご利用者|令和|日付|分$)/;
  let currentStaff = '';
  let helperLabelY = null;
  let helperLabelPage = null;

  // Step 1: 「ヘルパー名:」ラベルの行から、その後ろのテキストを抽出
  // PDFでは「ヘルパー名： 山田 愛美」のように1行に統合されているケースが多い
  for (const line of lines) {
    const m = line.text.match(/ヘルパー名\s*[::\uFF1A]\s*(.+)/);
    if (m) {
      const afterLabel = m[1].trim();
      // 「山田 愛美」「丁 翔」「中久木 匠」等の姓名パターンを抽出
      const nameMatch = afterLabel.match(/^([\u4E00-\u9FFF\u30A0-\u30FF]{1,3}[\s ]+[\u4E00-\u9FFF\u30A0-\u30FF]{1,3})/);
      if (nameMatch && !excludeStaff.test(nameMatch[1])) {
        currentStaff = nameMatch[1].replace(/\s+/g, ' ').trim();
        debug.notes.push(`スタッフ名 (ヘルパー名ラベル直後から検出): ${currentStaff}`);
        break;
      }
      // 同行に名前がなかった場合は、別行(近傍Y座標)からの検出にフォールバック
      helperLabelY = line.y;
      helperLabelPage = line.page;
      break;
    }
  }

  // Step 2: ラベル行に名前が含まれていなかった場合、近傍Y座標から探す(フォールバック)
  if (!currentStaff && helperLabelY !== null) {
    const candidates = lines
      .filter(l => l.page === helperLabelPage && Math.abs(l.y - helperLabelY) < 30)
      .filter(l => !/ヘルパー名/.test(l.text));
    for (const c of candidates) {
      const t = c.text.trim();
      const m = t.match(/^([\u4E00-\u9FFF\u30A0-\u30FF]{1,3}[\s ]+[\u4E00-\u9FFF\u30A0-\u30FF]{1,3})$/);
      if (m && !excludeStaff.test(t)) {
        currentStaff = t.replace(/\s+/g, ' ').trim();
        debug.notes.push(`スタッフ名 (ラベル近傍行から検出): ${currentStaff}`);
        break;
      }
    }
  }

  // フォールバック: ファイル名から
  if (!currentStaff) {
    const fnameMatch = filename.replace(/\.pdf$/i, '').match(/^([\u4E00-\u9FFF\u30A0-\u30FF]{2,8})/);
    if (fnameMatch) {
      currentStaff = fnameMatch[1];
      debug.notes.push(`スタッフ名 (ファイル名から): ${currentStaff}`);
    } else {
      debug.notes.push('スタッフ名が検出できませんでした');
    }
  }

  // データ行抽出
  // 行パターン: 5月1日(金) 07:00 ～ 07:25 25 大川 富子様 身体介護１・夜
  // 全角チルダ「～」(U+FF5E), 半角チルダ「~」, 波ダッシュ「〜」(U+301C) 全対応
  const lineRe = /(\d{1,2})月(\d{1,2})日\([^)]+\)\s+(\d{1,2}):(\d{2})\s*[~\uFF5E\u301C]\s*(\d{1,2}):(\d{2})\s+(\d+)\s+(.+)$/;

  let skipped = { office: 0, parseError: 0 };

  for (const line of lines) {
    const text = line.text.replace(/\s+/g, ' ').trim();
    if (!text) continue;
    debug.lines.push({ page: line.page, y: line.y, text });

    const m = text.match(lineRe);
    if (!m) continue;

    const month = parseInt(m[1], 10);
    const day = parseInt(m[2], 10);
    const sH = parseInt(m[3], 10);
    const sM = parseInt(m[4], 10);
    const eH = parseInt(m[5], 10);
    const eM = parseInt(m[6], 10);
    const rest = m[8].trim();

    // 「事務所」行は除外(スキップ)
    if (/^事務所/.test(rest)) {
      skipped.office++;
      continue;
    }

    // 利用者名とサービス内容を分離
    // 「大川 富子様 身体介護１・夜」 → 利用者「大川 富子」 + サービス「身体介護１・夜」
    let user = '';
    let svc = '';
    const userMatch = rest.match(/^(.+?)様\s*(.*)$/);
    if (userMatch) {
      user = userMatch[1].trim();
      svc = userMatch[2].trim();
    } else {
      // 「様」がない場合はスペース区切りで前半=利用者、後半=サービス
      const parts = rest.split(/\s+/);
      if (parts.length >= 3) {
        user = parts.slice(0, 2).join(' ');
        svc = parts.slice(2).join(' ');
      } else {
        skipped.parseError++;
        continue;
      }
    }

    const sMin = sH * 60 + sM;
    const eMin = eH * 60 + eM;
    if (eMin <= sMin) continue;

    out.push({
      date: `${month}/${day}`,
      staff: currentStaff || '(不明)',
      user,
      startSlot: Math.floor(sMin / 15),
      endSlot: Math.ceil(eMin / 15),
      service: svc,
      serviceClass: classifyService(svc),
              code: toServiceCode(svc)
    });
  }

  if (skipped.office > 0) debug.notes.push(`「事務所」行をスキップ: ${skipped.office}件`);
  if (skipped.parseError > 0) debug.notes.push(`パース失敗: ${skipped.parseError}件`);

  return { rows: out, detectedStaff: currentStaff, debug };
}

function classifyService(svc) {
  if (!svc) return 'その他';
  const s = svc.toString();
  if (/重訪|重度訪問/.test(s)) return '重訪';
  if (/移動支援|移動/.test(s)) return '移動';
  if (/家事/.test(s)) return '家事';
  if (/身体.*生活|複合/.test(s)) return '身生';
  if (/身体/.test(s)) return '身体';
  if (/生活/.test(s)) return '生活';
  if (/通院|乗降/.test(s)) return '通院';
  return 'その他';
}

// サービス名 → 短縮コード変換(Excel出力&アプリ表のサービスコード行用)
function toServiceCode(svc) {
  if (!svc) return '';
  const s = svc.toString();
  // 身体介護(夜間含む) → 身１
  if (/身体介護[1１]/.test(s)) return '身１';
  if (/身体介護[22]/.test(s)) return '身２';
  if (/身体介護[33]/.test(s)) return '身３';
  // 生活援助
  if (/生活援助[11]/.test(s)) return '生１';
  if (/生活援助[22]/.test(s)) return '生２';
  if (/生活援助[33]/.test(s)) return '生３';
  // 通院等乗降介助
  if (/通院.*乗降|乗降.*介助/.test(s)) return '通１';
  // 障害福祉系
  if (/重訪|重度訪問/.test(s)) return '重訪Ⅱ';
  if (/移動支援|移動/.test(s)) return '移動';
  if (/家事/.test(s)) return '家事';
  // 身体生活複合
  if (/身体.*生活/.test(s)) return '身生';
  // フォールバック: 身体/生活単独
  if (/^身体$/.test(s.trim())) return '身体';
  if (/^生活$/.test(s.trim())) return '生活';
  return s.substring(0, 4); // 4文字以内で省略
}

const SERVICE_COLORS = {
  '身体': '#CC99FF',
  '生活': '#60a5fa',
  '身生': '#c084fc',
  '通院': '#fbbf24',
  '移動': '#34d399',
  '家事': '#f472b6',
  '重訪': '#a78bfa',
  'その他': '#94a3b8'
};

// ============ 発光ボーダー付き要素 ============
function GlowFrame({ children, active, intense, style, ...props }) {
  const glowColor = '251, 191, 36';
  const baseAlpha = active ? 0.45 : 0.18;
  const intensity = intense ? 1.4 : 1;

  return (
    <div
      style={{
        position: 'relative',
        background: 'rgba(20, 23, 33, 0.6)',
        border: `1px solid rgba(${glowColor}, ${baseAlpha + 0.15})`,
        borderRadius: 6,
        boxShadow: `
          0 0 ${12 * intensity}px rgba(${glowColor}, ${baseAlpha * 0.5}),
          0 0 ${28 * intensity}px rgba(${glowColor}, ${baseAlpha * 0.25}),
          inset 0 0 ${8 * intensity}px rgba(${glowColor}, ${baseAlpha * 0.1})
        `,
        transition: 'all 0.4s cubic-bezier(0.4, 0, 0.2, 1)',
        ...style
      }}
      {...props}
    >
      {children}
      {/* 床面の光漏れ */}
      <div
        style={{
          position: 'absolute',
          bottom: -20,
          left: '10%',
          right: '10%',
          height: 20,
          background: `radial-gradient(ellipse at center top, rgba(${glowColor}, ${baseAlpha * 0.4}) 0%, transparent 70%)`,
          pointerEvents: 'none',
          opacity: active ? 1 : 0.5,
          transition: 'opacity 0.4s'
        }}
      />
    </div>
  );
}

// ============ メインアプリ ============
export default function StaffSchedule({ onBack }) {
  const [staffText, setStaffText] = useState('');
  const [staffFocused, setStaffFocused] = useState(false);
  const [parsedRows, setParsedRows] = useState([]);
  const [detectedStaff, setDetectedStaff] = useState([]); // PDF/CSVから検出されたスタッフ名一覧
  const [selectedStaff, setSelectedStaff] = useState(new Set()); // チェックされている検出スタッフ
  const [staffShifts, setStaffShifts] = useState({}); // {[staffName]: {start: '7:00', end: '21:00'}}
  const [staffTypes, setStaffTypes] = useState({}); // {[staffName]: '社員' | '登録'}
  const [copyStatus, setCopyStatus] = useState(''); // クリップボードコピーのステータス
  const [allUsers, setAllUsers] = useState([]);
  const [allDates, setAllDates] = useState([]);
  const [selectedUsers, setSelectedUsers] = useState(new Set());
  const [selectedDate, setSelectedDate] = useState('all');
  const [status, setStatus] = useState({ msg: '', error: false });
  const [hoveredCell, setHoveredCell] = useState(null);
  const [fileButtonHover, setFileButtonHover] = useState(false);
  const [activeButton, setActiveButton] = useState(null);
  const fileInputRef = useRef(null);

  // 登録済スタッフ名 (localStorage から1回読み込み)
  const [registeredNames, setRegisteredNames] = useState(
    () => new Set(loadStaff().map((s) => s.name))
  );

  // 出勤スタッフ = 手入力 + 検出スタッフから選択された人(重複排除)
  const workingStaff = useMemo(() => {
    const fromText = staffText.split(/[\n,、]/).map(s => s.trim()).filter(Boolean);
    const fromSelected = [...selectedStaff];
    return [...new Set([...fromText, ...fromSelected])];
  }, [staffText, selectedStaff]);

  const [debugInfo, setDebugInfo] = useState(null);
  const [showDebug, setShowDebug] = useState(false);

  // CSV / PDF 読み込み(両対応・複数ファイル可)
  // クリップボードに表をコピー(Excel/Google Sheets に貼り付け可能)
  async function handleCopyToClipboard() {
    if (!grid || workingStaff.length === 0) {
      alert('コピーするデータがありません。スタッフを選択してください。');
      return;
    }
    setCopyStatus('コピー中...');
    try {
      const tsv = generateTSV({
        workingStaff,
        staffShifts,
        staffTypes,
        grid: grid.grid,
        selectedDate
      });
      const ok = await copyTextToClipboard(tsv);
      if (ok) {
        setCopyStatus('✓ コピー完了。原本のA1セルを選択して Ctrl+V で貼り付けてください');
      } else {
        setCopyStatus('✗ コピーに失敗しました(ブラウザ権限を確認)');
      }
      setTimeout(() => setCopyStatus(''), 5000);
    } catch (err) {
      setCopyStatus('✗ エラー: ' + err.message);
      console.error(err);
    }
  }

  // Excel(.xlsx) ダウンロード: ひな形の書式を保持したまま値を流し込む
  async function handleDownloadXlsx() {
    if (!grid || workingStaff.length === 0) {
      alert('ダウンロードするデータがありません。スタッフを選択してください。');
      return;
    }
    setCopyStatus('Excel生成中...');
    try {
      await generateXlsx({
        workingStaff,
        staffShifts,
        staffTypes,
        grid: grid.grid,
        selectedDate
      });
      setCopyStatus('✓ Excelをダウンロードしました');
      setTimeout(() => setCopyStatus(''), 5000);
    } catch (err) {
      setCopyStatus('✗ エラー: ' + err.message);
      console.error(err);
    }
  }

  async function handleFile(e) {
    const files = [...e.target.files];
    if (!files.length) return;
    setStatus({ msg: '読み込み中...', error: false });
    setDebugInfo(null);

    try {
      const accumulated = [];
      const userSet = new Set();
      const dateSet = new Set();
      const sources = [];
      const debugBundles = [];

      for (const file of files) {
        const ext = file.name.toLowerCase().split('.').pop();

        if (ext === 'csv') {
          // ----- CSV -----
          const text = await file.text();
          const rows = parseCSV(text);
          if (rows.length < 2) {
            sources.push(`${file.name}: データ行なし`);
            continue;
          }
          const header = rows[0];
          const cols = detectColumns(header);
          if (cols.staff === -1 || cols.user === -1) {
            sources.push(`${file.name}: スタッフ/利用者列を検出できず`);
            continue;
          }
          let added = 0;
          for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            const staff = (r[cols.staff] || '').trim();
            const user = (r[cols.user] || '').trim();
            if (!staff || !user) continue;
            const date = cols.date >= 0 ? (r[cols.date] || '').trim() : '';
            const startMin = cols.startTime >= 0 ? parseTime(r[cols.startTime]) : null;
            const endMin = cols.endTime >= 0 ? parseTime(r[cols.endTime]) : null;
            if (startMin == null || endMin == null) continue;
            const svc = cols.service >= 0 ? (r[cols.service] || '').trim() : '';
            accumulated.push({
              date, staff, user,
              startSlot: Math.floor(startMin / 15),
              endSlot: Math.ceil((endMin === 0 ? 1440 : endMin) / 15),
              service: svc,
              serviceClass: classifyService(svc),
              code: toServiceCode(svc)
            });
            userSet.add(user);
            if (date) dateSet.add(date);
            added++;
          }
          sources.push(`${file.name}: CSV ${added}件`);

        } else if (ext === 'pdf') {
          // ----- PDF -----
          const lines = await pdfToLines(file);
          const { rows: pdfRows, detectedStaff, debug } = extractFromPdfLines(lines, file.name);
          for (const r of pdfRows) {
            accumulated.push(r);
            userSet.add(r.user);
            if (r.date) dateSet.add(r.date);
          }
          sources.push(`${file.name}: PDF ${pdfRows.length}件 (スタッフ=${detectedStaff || '未検出'})`);
          debugBundles.push({ file: file.name, staff: detectedStaff, debug, rowCount: pdfRows.length });

        } else {
          sources.push(`${file.name}: 未対応形式`);
        }
      }

      if (accumulated.length === 0) {
        if (debugBundles.length) setDebugInfo(debugBundles);
        throw new Error('有効なサービス実績データがありません\n' + sources.join('\n') + '\n\n下部「PDF読取結果を確認」でPDFから読み取れた行を確認できます');
      }

      // 成功時もPDFがあればdebug保持
      if (debugBundles.length) setDebugInfo(debugBundles);

      // 検出スタッフ一覧を集約 + デフォルト全選択
      const staffSet = new Set();
      for (const r of accumulated) {
        if (r.staff && r.staff !== '(不明)') staffSet.add(r.staff);
      }
      const detectedList = [...staffSet].sort();
      setDetectedStaff(detectedList);
      setSelectedStaff(new Set(detectedList));

      // 登録済みスタッフから区分を自動補完
      const autoTypes = {};
      for (const name of detectedList) {
        const t = lookupType(name);
        if (t) autoTypes[name] = t;
      }
      if (Object.keys(autoTypes).length > 0) {
        setStaffTypes((prev) => ({ ...autoTypes, ...prev }));
      }

      const users = [...userSet].sort();
      const dates = [...dateSet].sort();
      setParsedRows(accumulated);
      setAllUsers(users);
      setAllDates(dates);
      setSelectedUsers(new Set(users));
      setSelectedDate(dates.length > 1 ? dates[0] : 'all');
      setStatus({
        msg: `✓ 合計 ${accumulated.length}件 読み込み\n` + sources.join('\n'),
        error: false
      });
    } catch (err) {
      setStatus({ msg: '✗ ' + err.message, error: true });
      setParsedRows([]);
    } finally {
      // 同じファイルを連続選択できるようにリセット
      e.target.value = '';
    }
  }

  // グリッド構築
  const grid = useMemo(() => {
    if (workingStaff.length === 0 || parsedRows.length === 0) return null;
    const filtered = parsedRows.filter(r =>
      workingStaff.includes(r.staff) &&
      selectedUsers.has(r.user) &&
      (selectedDate === 'all' || r.date === selectedDate)
    );
    const g = {};
    workingStaff.forEach(s => { g[s] = Array(96).fill(null); });
    filtered.forEach((r, idx) => {
      if (!g[r.staff]) return;
      for (let i = r.startSlot; i < r.endSlot && i < 96; i++) {
        const cell = g[r.staff][i];
        const entry = {
          user: r.user,
          serviceClass: r.serviceClass,
          service: r.service,
          code: r.code,
          groupId: `${r.staff}-${idx}`, // 同一サービス連続スロット識別用
          isStart: i === r.startSlot,
          isEnd: i === r.endSlot - 1,
          slotsInGroup: r.endSlot - r.startSlot
        };
        if (cell) {
          cell.user = cell.user.includes(r.user) ? cell.user : cell.user + ',' + r.user;
        } else {
          g[r.staff][i] = entry;
        }
      }
    });

    // === アラート①: 選択日にサービスがあるのに表に出てない利用者 ===
    // (出勤者の漏れ、または出勤者外への割当の可能性)
    const usersShownInTable = new Set(filtered.map(r => r.user));
    // 「選択日にサービスがある利用者」全体
    const usersWithServiceOnDate = new Set(
      parsedRows
        .filter(r => selectedDate === 'all' || r.date === selectedDate)
        .filter(r => selectedUsers.has(r.user))
        .map(r => r.user)
    );
    // 漏れ = サービスはあるが表に出てない
    const missingUsers = [...usersWithServiceOnDate]
      .filter(u => !usersShownInTable.has(u))
      .sort();

    // 漏れた利用者ごとに、出勤者外で担当しているスタッフを集計
    const missingDetails = missingUsers.map(u => {
      const recordsOnDate = parsedRows.filter(r =>
        r.user === u &&
        (selectedDate === 'all' || r.date === selectedDate)
      );
      const inOtherStaff = [...new Set(recordsOnDate.map(r => r.staff))]
        .filter(s => !workingStaff.includes(s));
      return { user: u, otherStaff: inOtherStaff, totalRecords: recordsOnDate.length };
    });

    // === アラート②: 選択日にサービス予定があるのに出勤チェックがついてないスタッフ ===
    // ただし「選択されている利用者へのサービス」を持つスタッフだけが対象
    // (=利用者フィルタで外している利用者しか担当していないスタッフはアラート対象外)
    const staffWithServiceOnDate = new Set(
      parsedRows
        .filter(r => selectedDate === 'all' || r.date === selectedDate)
        .filter(r => selectedUsers.has(r.user))  // ★選択されている利用者のサービスだけ
        .map(r => r.staff)
    );
    const unmarkedStaff = [...staffWithServiceOnDate]
      .filter(s => !workingStaff.includes(s))
      .sort();
    // 各スタッフが担当している「選択中の利用者」のみ集計
    const unmarkedStaffDetails = unmarkedStaff.map(s => {
      const recordsOnDate = parsedRows.filter(r =>
        r.staff === s &&
        selectedUsers.has(r.user) &&  // ★選択されている利用者のみ
        (selectedDate === 'all' || r.date === selectedDate)
      );
      const users = [...new Set(recordsOnDate.map(r => r.user))];
      return { staff: s, userCount: users.length, recordCount: recordsOnDate.length, users };
    });

    return { grid: g, count: filtered.length, missingDetails, unmarkedStaffDetails };
  }, [workingStaff, parsedRows, selectedUsers, selectedDate]);

  const toggleUser = (u) => {
    const next = new Set(selectedUsers);
    if (next.has(u)) next.delete(u); else next.add(u);
    setSelectedUsers(next);
  };

  const baseStyle = {
    minHeight: '100vh',
    background: 'radial-gradient(ellipse at top, #0f1419 0%, #050709 100%)',
    color: '#e5e7eb',
    fontFamily: '"Hiragino Sans", "Yu Gothic", "Noto Sans JP", sans-serif',
    padding: 24,
    boxSizing: 'border-box'
  };

  return (
    <div style={baseStyle}>
      {/* 月のアクセント */}
      <div style={{
        position: 'fixed', top: 30, right: 60,
        width: 40, height: 40,
        background: 'radial-gradient(circle, #fef3c7 0%, #fbbf24 60%, transparent 70%)',
        borderRadius: '50%',
        opacity: 0.6,
        clipPath: 'polygon(40% 0%, 100% 0%, 100% 100%, 40% 100%, 70% 70%, 70% 30%)',
        pointerEvents: 'none'
      }} />

      <header style={{ marginBottom: 32, position: 'relative' }}>
        {onBack && (
          <button
            onClick={onBack}
            style={{
              position: 'absolute',
              top: -4,
              left: 0,
              padding: '6px 12px',
              background: 'rgba(0,0,0,0.4)',
              border: '1px solid rgba(251, 191, 36, 0.3)',
              color: '#fef3c7',
              fontFamily: 'inherit',
              fontSize: 11,
              letterSpacing: '0.15em',
              borderRadius: 3,
              cursor: 'pointer'
            }}
          >← ダッシュボード</button>
        )}
        <h1 style={{
          fontSize: 26,
          fontWeight: 300,
          letterSpacing: '0.15em',
          margin: 0,
          marginTop: onBack ? 28 : 0,
          color: '#fef3c7',
          textShadow: '0 0 20px rgba(251, 191, 36, 0.4)'
        }}>
          スタッフ・利用者 対応表
        </h1>
        <p style={{
          fontSize: 11,
          color: 'rgba(251, 191, 36, 0.5)',
          letterSpacing: '0.3em',
          marginTop: 6,
          textTransform: 'uppercase'
        }}>
          Kaipoke ─ Visit Care Schedule
        </p>
      </header>

      <div style={{
        display: 'grid',
        gridTemplateColumns: 'minmax(340px, 380px) 1fr',
        gap: 28,
        alignItems: 'start'
      }}>
        {/* ===== 左パネル ===== */}
        <aside style={{ display: 'flex', flexDirection: 'column', gap: 28 }}>
          {/* 01: CSV/PDFアップロード */}
          <GlowFrame active={fileButtonHover || activeButton === 'file'} intense={activeButton === 'file'}>
            <div style={{ padding: 18 }}>
              <SectionLabel num="01" label="CSV / PDF アップロード" />
              <input
                ref={fileInputRef}
                type="file"
                accept=".csv,.pdf,text/csv,application/pdf"
                multiple
                onChange={handleFile}
                style={{ display: 'none' }}
              />
              <button
                onClick={() => {
                  setActiveButton('file');
                  fileInputRef.current?.click();
                  setTimeout(() => setActiveButton(null), 600);
                }}
                onMouseEnter={() => setFileButtonHover(true)}
                onMouseLeave={() => setFileButtonHover(false)}
                style={{
                  width: '100%',
                  padding: '12px',
                  background: fileButtonHover ? 'rgba(251, 191, 36, 0.12)' : 'rgba(0,0,0,0.3)',
                  border: `1px solid rgba(251, 191, 36, ${fileButtonHover ? 0.5 : 0.25})`,
                  borderRadius: 4,
                  color: '#fef3c7',
                  fontFamily: 'inherit',
                  fontSize: 12,
                  letterSpacing: '0.1em',
                  cursor: 'pointer',
                  transition: 'all 0.3s'
                }}
              >
                ファイルを選択
              </button>
              <p style={{
                marginTop: 8,
                fontSize: 10,
                color: 'rgba(251, 191, 36, 0.4)',
                letterSpacing: '0.05em',
                lineHeight: 1.5
              }}>
                カイポケCSV / スタッフ別PDF (複数可)
              </p>
              {status.msg && (
                <div style={{
                  marginTop: 12,
                  padding: '8px 10px',
                  fontSize: 11,
                  color: status.error ? '#fca5a5' : '#fde68a',
                  background: status.error ? 'rgba(239, 68, 68, 0.1)' : 'rgba(251, 191, 36, 0.08)',
                  border: `1px solid ${status.error ? 'rgba(239, 68, 68, 0.3)' : 'rgba(251, 191, 36, 0.2)'}`,
                  borderRadius: 3,
                  whiteSpace: 'pre-wrap',
                  lineHeight: 1.5
                }}>
                  {status.msg}
                </div>
              )}

              {debugInfo && (
                <button
                  onClick={() => setShowDebug(v => !v)}
                  style={{
                    marginTop: 10,
                    width: '100%',
                    padding: '6px',
                    background: 'rgba(0,0,0,0.3)',
                    border: '1px solid rgba(251, 191, 36, 0.25)',
                    color: 'rgba(251, 191, 36, 0.7)',
                    fontFamily: 'inherit',
                    fontSize: 11,
                    borderRadius: 3,
                    cursor: 'pointer',
                    letterSpacing: '0.05em'
                  }}
                >
                  {showDebug ? 'PDF読取結果を閉じる' : 'PDF読取結果を確認'}
                </button>
              )}

              {allDates.length > 1 && (
                <select
                  value={selectedDate}
                  onChange={(e) => setSelectedDate(e.target.value)}
                  style={{
                    marginTop: 12,
                    width: '100%',
                    padding: '8px',
                    background: 'rgba(0,0,0,0.4)',
                    border: '1px solid rgba(251, 191, 36, 0.2)',
                    color: '#fef3c7',
                    fontFamily: 'inherit',
                    fontSize: 12,
                    borderRadius: 3,
                    outline: 'none'
                  }}
                >
                  <option value="all">全期間まとめて表示</option>
                  {allDates.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
              )}
            </div>
          </GlowFrame>

          {/* 02: 出勤スタッフ選択 */}
          <GlowFrame active={detectedStaff.length > 0} intense={selectedStaff.size > 0}>
            <div style={{ padding: 18 }}>
              <SectionLabel num="02" label="出勤スタッフ" />

              {detectedStaff.length === 0 ? (
                <div style={{
                  padding: '20px 10px',
                  textAlign: 'center',
                  fontSize: 11,
                  color: 'rgba(251, 191, 36, 0.4)',
                  letterSpacing: '0.05em',
                  lineHeight: 1.7
                }}>
                  上のファイルをアップロードすると<br />
                  検出されたスタッフがここに表示されます
                </div>
              ) : (
                <>
                  <div style={{
                    marginBottom: 10,
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    fontSize: 10,
                    color: 'rgba(251, 191, 36, 0.6)',
                    letterSpacing: '0.1em'
                  }}>
                    <span>検出: {detectedStaff.length}名 / 選択: {selectedStaff.size}名</span>
                    <div style={{ display: 'flex', gap: 6 }}>
                      <button
                        onClick={() => setSelectedStaff(new Set(detectedStaff))}
                        style={{
                          padding: '3px 10px',
                          fontSize: 10,
                          background: 'rgba(0,0,0,0.3)',
                          border: '1px solid rgba(251, 191, 36, 0.3)',
                          color: '#fef3c7',
                          borderRadius: 2,
                          cursor: 'pointer',
                          fontFamily: 'inherit',
                          letterSpacing: '0.1em'
                        }}
                      >全選択</button>
                      <button
                        onClick={() => setSelectedStaff(new Set())}
                        style={{
                          padding: '3px 10px',
                          fontSize: 10,
                          background: 'rgba(0,0,0,0.3)',
                          border: '1px solid rgba(251, 191, 36, 0.3)',
                          color: '#fef3c7',
                          borderRadius: 2,
                          cursor: 'pointer',
                          fontFamily: 'inherit',
                          letterSpacing: '0.1em'
                        }}
                      >全解除</button>
                    </div>
                  </div>
                  <div style={{
                    maxHeight: 360,
                    overflowY: 'auto',
                    background: 'rgba(0,0,0,0.3)',
                    border: '1px solid rgba(251, 191, 36, 0.15)',
                    borderRadius: 3,
                    padding: 4
                  }}>
                    {detectedStaff.map(s => (
                      <DetectedStaffItem
                        key={s}
                        name={s}
                        checked={selectedStaff.has(s)}
                        registered={registeredNames.has(s)}
                        onRegister={() => {
                          const currentType = staffTypes[s] || '社員';
                          addStaff(s, currentType);
                          setRegisteredNames(prev => {
                            const next = new Set(prev);
                            next.add(s);
                            return next;
                          });
                        }}
                        shift={staffShifts[s] || { start: '07:00', end: '21:00' }}
                        type={staffTypes[s] || '社員'}
                        onShiftChange={(newShift) => {
                          setStaffShifts(prev => ({ ...prev, [s]: newShift }));
                        }}
                        onTypeChange={(newType) => {
                          setStaffTypes(prev => ({ ...prev, [s]: newType }));
                        }}
                        onToggle={() => {
                          setSelectedStaff(prev => {
                            const next = new Set(prev);
                            if (next.has(s)) next.delete(s); else next.add(s);
                            return next;
                          });
                          // 初回チェック時にデフォルト勤務時間/区分を設定
                          setStaffShifts(prev => {
                            if (prev[s]) return prev;
                            return { ...prev, [s]: { start: '07:00', end: '21:00' } };
                          });
                          setStaffTypes(prev => {
                            if (prev[s]) return prev;
                            return { ...prev, [s]: '社員' };
                          });
                        }}
                      />
                    ))}
                  </div>

                  <details style={{
                    marginTop: 14,
                    fontSize: 11,
                    color: 'rgba(251, 191, 36, 0.5)'
                  }}>
                    <summary style={{
                      cursor: 'pointer',
                      letterSpacing: '0.1em',
                      padding: '4px 0'
                    }}>
                      手入力で追加
                    </summary>
                    <textarea
                      value={staffText}
                      onChange={(e) => setStaffText(e.target.value)}
                      onFocus={() => setStaffFocused(true)}
                      onBlur={() => setStaffFocused(false)}
                      placeholder="改行 or カンマ区切り"
                      style={{
                        marginTop: 6,
                        width: '100%',
                        minHeight: 50,
                        background: 'rgba(0,0,0,0.4)',
                        border: '1px solid rgba(251, 191, 36, 0.15)',
                        borderRadius: 4,
                        color: '#fef3c7',
                        fontFamily: 'inherit',
                        fontSize: 12,
                        padding: 8,
                        outline: 'none',
                        resize: 'vertical',
                        boxSizing: 'border-box'
                      }}
                    />
                  </details>
                </>
              )}
            </div>
          </GlowFrame>

          {/* 利用者フィルタ */}
          {allUsers.length > 0 && (
            <GlowFrame>
              <div style={{ padding: 18 }}>
                <SectionLabel num="03" label="利用者で絞り込み" />
                <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
                  <SmallButton
                    onClick={() => setSelectedUsers(new Set(allUsers))}
                    active={activeButton === 'all'}
                    onPress={() => { setActiveButton('all'); setTimeout(()=>setActiveButton(null), 500); }}
                  >全選択</SmallButton>
                  <SmallButton
                    onClick={() => setSelectedUsers(new Set())}
                    active={activeButton === 'none'}
                    onPress={() => { setActiveButton('none'); setTimeout(()=>setActiveButton(null), 500); }}
                  >全解除</SmallButton>
                </div>
                <div style={{
                  maxHeight: 240,
                  overflowY: 'auto',
                  background: 'rgba(0,0,0,0.3)',
                  border: '1px solid rgba(251, 191, 36, 0.1)',
                  borderRadius: 3,
                  padding: 6
                }}>
                  {allUsers.map(u => (
                    <UserCheckbox
                      key={u}
                      user={u}
                      checked={selectedUsers.has(u)}
                      onToggle={() => toggleUser(u)}
                    />
                  ))}
                </div>
              </div>
            </GlowFrame>
          )}

          {/* 04: コピーで原本に貼り付け */}
          {grid && workingStaff.length > 0 && (
            <GlowFrame active={!!copyStatus.startsWith('✓')} intense={!!copyStatus.startsWith('✓')}>
              <div style={{ padding: 18 }}>
                <SectionLabel num="04" label="原本にコピペ" />
                <p style={{
                  fontSize: 10,
                  color: 'rgba(251, 191, 36, 0.5)',
                  letterSpacing: '0.05em',
                  lineHeight: 1.5,
                  marginBottom: 12
                }}>
                  {selectedDate !== 'all' ? selectedDate : '全期間'}のデータを<br/>
                  タブ区切り(TSV)で出力 → 原本に貼付
                </p>
                <button
                  onClick={handleCopyToClipboard}
                  style={{
                    width: '100%',
                    padding: '14px',
                    background: 'rgba(251, 191, 36, 0.15)',
                    border: '1px solid rgba(251, 191, 36, 0.6)',
                    borderRadius: 4,
                    color: '#fef3c7',
                    fontFamily: 'inherit',
                    fontSize: 13,
                    fontWeight: 500,
                    letterSpacing: '0.15em',
                    cursor: 'pointer',
                    transition: 'all 0.3s',
                    boxShadow: '0 0 8px rgba(251, 191, 36, 0.2)'
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.background = 'rgba(251, 191, 36, 0.3)';
                    e.currentTarget.style.boxShadow = '0 0 16px rgba(251, 191, 36, 0.5)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.background = 'rgba(251, 191, 36, 0.15)';
                    e.currentTarget.style.boxShadow = '0 0 8px rgba(251, 191, 36, 0.2)';
                  }}
                >
                  📋 TSVをコピー
                </button>
                <button
                  onClick={handleDownloadXlsx}
                  style={{
                    width: '100%',
                    marginTop: 8,
                    padding: '14px',
                    background: 'rgba(34, 197, 94, 0.15)',
                    border: '1px solid rgba(34, 197, 94, 0.6)',
                    borderRadius: 4,
                    color: '#bbf7d0',
                    fontFamily: 'inherit',
                    fontSize: 13,
                    fontWeight: 500,
                    letterSpacing: '0.15em',
                    cursor: 'pointer',
                    transition: 'all 0.3s',
                    boxShadow: '0 0 8px rgba(34, 197, 94, 0.2)'
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.background = 'rgba(34, 197, 94, 0.3)';
                    e.currentTarget.style.boxShadow = '0 0 16px rgba(34, 197, 94, 0.5)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.background = 'rgba(34, 197, 94, 0.15)';
                    e.currentTarget.style.boxShadow = '0 0 8px rgba(34, 197, 94, 0.2)';
                  }}
                >
                  ⬇ Excelダウンロード（書式保持）
                </button>
                {copyStatus && (
                  <div style={{
                    marginTop: 10,
                    padding: '8px 10px',
                    fontSize: 11,
                    color: copyStatus.startsWith('✗')
                      ? '#fca5a5'
                      : (copyStatus.startsWith('✓') ? '#86efac' : '#fde68a'),
                    background: copyStatus.startsWith('✗')
                      ? 'rgba(239, 68, 68, 0.1)'
                      : (copyStatus.startsWith('✓') ? 'rgba(34, 197, 94, 0.1)' : 'rgba(251, 191, 36, 0.08)'),
                    border: `1px solid ${copyStatus.startsWith('✗') ? 'rgba(239, 68, 68, 0.3)' : (copyStatus.startsWith('✓') ? 'rgba(34, 197, 94, 0.3)' : 'rgba(251, 191, 36, 0.2)')}`,
                    borderRadius: 3,
                    lineHeight: 1.5
                  }}>
                    {copyStatus}
                  </div>
                )}
                <div style={{
                  marginTop: 14,
                  paddingTop: 12,
                  borderTop: '1px solid rgba(251, 191, 36, 0.1)',
                  fontSize: 10,
                  color: 'rgba(251, 191, 36, 0.4)',
                  lineHeight: 1.7
                }}>
                  使い方:<br/>
                  ① ボタンを押す<br/>
                  ② テンプレ原本(Excel/スプレッドシート)を開く<br/>
                  ③ <strong>A1セル</strong>を選択<br/>
                  ④ Ctrl+V(または右クリック→貼り付け)
                </div>
              </div>
            </GlowFrame>
          )}
        </aside>

        {/* ===== 右メインエリア ===== */}
        <main>
          {showDebug && debugInfo ? (
            <DebugPanel bundles={debugInfo} />
          ) : !grid ? (
            <GlowFrame style={{ minHeight: 400 }}>
              <div style={{
                padding: 80,
                textAlign: 'center',
                color: 'rgba(251, 191, 36, 0.4)',
                fontSize: 13,
                letterSpacing: '0.1em'
              }}>
                <div style={{ fontSize: 40, marginBottom: 16, opacity: 0.5 }}>◯</div>
                {workingStaff.length === 0
                  ? (detectedStaff.length > 0
                      ? 'PDFから検出されたスタッフをチェックしてください'
                      : '出勤スタッフを入力 or CSV/PDFをアップロードしてください')
                  : 'CSV/PDFをアップロードしてください'}
              </div>
            </GlowFrame>
          ) : (
            <ScheduleTable
              grid={grid.grid}
              count={grid.count}
              staffList={workingStaff}
              userCount={selectedUsers.size}
              selectedDate={selectedDate}
              hoveredCell={hoveredCell}
              setHoveredCell={setHoveredCell}
              staffShifts={staffShifts}
              staffTypes={staffTypes}
              missingDetails={grid.missingDetails}
              unmarkedStaffDetails={grid.unmarkedStaffDetails}
            />
          )}
        </main>
      </div>
    </div>
  );
}

// ============ 補助コンポーネント ============
function SectionLabel({ num, label }) {
  return (
    <div style={{
      display: 'flex',
      alignItems: 'center',
      gap: 10,
      marginBottom: 14,
      paddingBottom: 10,
      borderBottom: '1px solid rgba(251, 191, 36, 0.1)'
    }}>
      <span style={{
        fontSize: 10,
        color: 'rgba(251, 191, 36, 0.5)',
        fontFamily: 'monospace',
        letterSpacing: '0.1em'
      }}>{num}</span>
      <span style={{
        fontSize: 11,
        color: '#fef3c7',
        letterSpacing: '0.2em',
        textTransform: 'uppercase'
      }}>{label}</span>
    </div>
  );
}

function SmallButton({ children, onClick, active, onPress }) {
  const [hover, setHover] = useState(false);
  const lit = hover || active;
  return (
    <button
      onClick={() => { onPress(); onClick(); }}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        flex: 1,
        padding: '6px 10px',
        fontSize: 11,
        background: lit ? 'rgba(251, 191, 36, 0.15)' : 'rgba(0,0,0,0.3)',
        border: `1px solid rgba(251, 191, 36, ${lit ? 0.5 : 0.2})`,
        color: '#fef3c7',
        fontFamily: 'inherit',
        borderRadius: 3,
        cursor: 'pointer',
        letterSpacing: '0.1em',
        transition: 'all 0.3s',
        boxShadow: active ? '0 0 16px rgba(251, 191, 36, 0.4)' : 'none'
      }}
    >
      {children}
    </button>
  );
}

function UserCheckbox({ user, checked, onToggle }) {
  const [hover, setHover] = useState(false);
  return (
    <div
      onClick={onToggle}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      role="checkbox"
      aria-checked={checked}
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === ' ' || e.key === 'Enter') {
          e.preventDefault();
          onToggle();
        }
      }}
      style={{
        display: 'flex',
        alignItems: 'center',
        padding: '6px 8px',
        fontSize: 12,
        cursor: 'pointer',
        background: hover ? 'rgba(251, 191, 36, 0.08)' : 'transparent',
        borderRadius: 2,
        transition: 'background 0.2s',
        color: checked ? '#fef3c7' : 'rgba(229, 231, 235, 0.5)',
        userSelect: 'none',
        outline: 'none'
      }}
    >
      <span style={{
        width: 12,
        height: 12,
        marginRight: 8,
        border: `1px solid rgba(251, 191, 36, ${checked ? 0.7 : 0.25})`,
        background: checked ? 'rgba(251, 191, 36, 0.85)' : 'transparent',
        borderRadius: 2,
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        boxShadow: checked ? '0 0 8px rgba(251, 191, 36, 0.5)' : 'none',
        transition: 'all 0.3s',
        flexShrink: 0,
        pointerEvents: 'none'
      }}>
        {checked && (
          <span style={{
            color: '#0a0d12',
            fontSize: 9,
            fontWeight: 'bold',
            lineHeight: 1
          }}>✓</span>
        )}
      </span>
      <span style={{ pointerEvents: 'none' }}>{user}</span>
    </div>
  );
}

function DetectedStaffItem({ name, checked, onToggle, shift, onShiftChange, type, onTypeChange, registered = true, onRegister }) {
  const [hover, setHover] = useState(false);
  return (
    <div
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        padding: '8px 10px',
        background: hover
          ? 'rgba(251, 191, 36, 0.12)'
          : (checked ? 'rgba(251, 191, 36, 0.05)' : 'transparent'),
        borderRadius: 3,
        transition: 'background 0.15s',
        color: checked ? '#fef3c7' : 'rgba(229, 231, 235, 0.5)',
        userSelect: 'none',
        marginBottom: 2
      }}
    >
      {/* 1段目: チェック + 名前 + 社員/登録 */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 10
        }}
      >
        <div
          onClick={onToggle}
          role="checkbox"
          aria-checked={checked}
          tabIndex={0}
          onKeyDown={(e) => {
            if (e.key === ' ' || e.key === 'Enter') {
              e.preventDefault();
              onToggle();
            }
          }}
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 10,
            cursor: 'pointer',
            outline: 'none',
            flex: 1,
            minWidth: 0
          }}
        >
          <span style={{
            width: 16,
            height: 16,
            border: `1.5px solid rgba(251, 191, 36, ${checked ? 0.9 : 0.35})`,
            background: checked ? '#fbbf24' : 'transparent',
            borderRadius: 3,
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: checked ? '0 0 12px rgba(251, 191, 36, 0.7)' : 'none',
            transition: 'all 0.25s',
            flexShrink: 0,
            pointerEvents: 'none'
          }}>
            {checked && (
              <span style={{
                color: '#0a0d12',
                fontSize: 12,
                fontWeight: 'bold',
                lineHeight: 1
              }}>✓</span>
            )}
          </span>
          <span style={{
            pointerEvents: 'none',
            fontSize: 14,
            fontWeight: 500,
            letterSpacing: '0.02em',
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis'
          }}>{name}</span>
        </div>

        {/* 社員/登録 トグル(チェック時のみ表示) */}
        {checked && (
          <div style={{
            display: 'flex',
            gap: 0,
            border: '1px solid rgba(251, 191, 36, 0.3)',
            borderRadius: 3,
            overflow: 'hidden',
            flexShrink: 0
          }}>
            <button
              onClick={(e) => { e.stopPropagation(); onTypeChange('社員'); }}
              style={{
                padding: '3px 9px',
                fontSize: 11,
                background: type === '社員' ? 'rgba(251, 191, 36, 0.8)' : 'transparent',
                color: type === '社員' ? '#0a0d12' : 'rgba(251, 191, 36, 0.7)',
                border: 'none',
                fontFamily: 'inherit',
                cursor: 'pointer',
                fontWeight: type === '社員' ? 600 : 400,
                transition: 'all 0.2s',
                boxShadow: type === '社員' ? '0 0 10px rgba(251, 191, 36, 0.5)' : 'none'
              }}
            >社員</button>
            <button
              onClick={(e) => { e.stopPropagation(); onTypeChange('登録'); }}
              style={{
                padding: '3px 9px',
                fontSize: 11,
                background: type === '登録' ? 'rgba(251, 191, 36, 0.8)' : 'transparent',
                color: type === '登録' ? '#0a0d12' : 'rgba(251, 191, 36, 0.7)',
                border: 'none',
                borderLeft: '1px solid rgba(251, 191, 36, 0.3)',
                fontFamily: 'inherit',
                cursor: 'pointer',
                fontWeight: type === '登録' ? 600 : 400,
                transition: 'all 0.2s',
                boxShadow: type === '登録' ? '0 0 10px rgba(251, 191, 36, 0.5)' : 'none'
              }}
            >登録</button>
          </div>
        )}
      </div>

      {/* 1.5段目: 未登録警告 + マスター追加ボタン (未登録かつチェック時) */}
      {checked && !registered && (
        <div style={{
          marginTop: 6,
          marginLeft: 26,
          display: 'flex',
          alignItems: 'center',
          gap: 8
        }}>
          <span
            style={{
              fontSize: 10,
              padding: '2px 6px',
              background: 'rgba(239, 68, 68, 0.15)',
              border: '1px solid rgba(239, 68, 68, 0.5)',
              color: '#fca5a5',
              borderRadius: 3,
              letterSpacing: '0.1em'
            }}
          >未登録</span>
          {onRegister && (
            <button
              onClick={(e) => { e.stopPropagation(); onRegister(); }}
              title={`「${name}」を現在の区分「${type}」でマスター追加`}
              style={{
                fontSize: 10,
                padding: '3px 10px',
                background: 'rgba(34, 197, 94, 0.15)',
                border: '1px solid rgba(34, 197, 94, 0.55)',
                color: '#bbf7d0',
                fontFamily: 'inherit',
                letterSpacing: '0.1em',
                borderRadius: 3,
                cursor: 'pointer'
              }}
            >＋マスター追加</button>
          )}
        </div>
      )}

      {/* 2段目: 時間入力欄(チェック時のみ表示) */}
      {checked && (
        <div style={{
          marginTop: 8,
          marginLeft: 26,
          display: 'flex',
          alignItems: 'center',
          gap: 8
        }}>
          <span style={{
            fontSize: 10,
            color: 'rgba(0, 255, 0, 0.5)',
            letterSpacing: '0.1em',
            marginRight: 2
          }}>勤務</span>
          <input
            type="time"
            value={shift.start}
            onChange={(e) => onShiftChange({ ...shift, start: e.target.value })}
            onClick={(e) => e.stopPropagation()}
            style={{
              width: 88,
              padding: '5px 6px',
              background: 'rgba(0, 0, 0, 0.5)',
              border: '1px solid rgba(0, 255, 0, 0.4)',
              color: '#00FF00',
              fontFamily: 'inherit',
              fontSize: 12,
              borderRadius: 3,
              outline: 'none',
              colorScheme: 'dark'
            }}
          />
          <span style={{
            fontSize: 11,
            color: 'rgba(0, 255, 0, 0.7)'
          }}>〜</span>
          <input
            type="time"
            value={shift.end}
            onChange={(e) => onShiftChange({ ...shift, end: e.target.value })}
            onClick={(e) => e.stopPropagation()}
            style={{
              width: 88,
              padding: '5px 6px',
              background: 'rgba(0, 0, 0, 0.5)',
              border: '1px solid rgba(0, 255, 0, 0.4)',
              color: '#00FF00',
              fontFamily: 'inherit',
              fontSize: 12,
              borderRadius: 3,
              outline: 'none',
              colorScheme: 'dark'
            }}
          />
        </div>
      )}
    </div>
  );
}

function ScheduleTable({ grid, count, staffList, userCount, selectedDate, hoveredCell, setHoveredCell, staffShifts, staffTypes, missingDetails, unmarkedStaffDetails }) {
  // 各スタッフの勤務時間スロット範囲を計算
  // shiftSlotMap[staffName] = { startSlot, endSlot } (15分単位、未入力なら null)
  const shiftSlotMap = {};
  for (const s of staffList) {
    const shift = staffShifts ? staffShifts[s] : null;
    if (!shift || !shift.start || !shift.end) { shiftSlotMap[s] = null; continue; }
    const [sh, sm] = shift.start.split(':').map(n => parseInt(n, 10));
    const [eh, em] = shift.end.split(':').map(n => parseInt(n, 10));
    if (isNaN(sh) || isNaN(sm) || isNaN(eh) || isNaN(em)) { shiftSlotMap[s] = null; continue; }
    const startSlot = Math.floor((sh * 60 + sm) / 15);
    const endSlot = Math.floor(((eh === 0 && em === 0 ? 1440 : eh * 60 + em)) / 15);
    shiftSlotMap[s] = { startSlot, endSlot };
  }

  return (
    <div>
      {/* サマリ */}
      <div style={{
        display: 'flex',
        gap: 24,
        marginBottom: 16,
        padding: '12px 18px',
        background: 'rgba(20, 23, 33, 0.4)',
        border: '1px solid rgba(251, 191, 36, 0.15)',
        borderRadius: 4,
        fontSize: 11,
        color: 'rgba(251, 191, 36, 0.6)',
        letterSpacing: '0.1em'
      }}>
        <SummaryItem label="STAFF" value={staffList.length} />
        <SummaryItem label="USERS" value={userCount} />
        <SummaryItem label="RECORDS" value={count} />
        {selectedDate !== 'all' && <SummaryItem label="DATE" value={selectedDate} />}
      </div>

      {/* ===== アラート①: 出勤未チェックのスタッフ ===== */}
      {unmarkedStaffDetails && unmarkedStaffDetails.length > 0 && (
        <div style={{
          marginBottom: 12,
          padding: '14px 18px',
          background: 'rgba(239, 68, 68, 0.12)',
          border: '1px solid rgba(239, 68, 68, 0.55)',
          borderRadius: 4,
          boxShadow: '0 0 22px rgba(239, 68, 68, 0.22), inset 0 0 12px rgba(239, 68, 68, 0.12)'
        }}>
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            marginBottom: 10,
            fontSize: 13,
            fontWeight: 600,
            color: '#fca5a5'
          }}>
            <span style={{
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: 20,
              height: 20,
              borderRadius: '50%',
              background: 'rgba(239, 68, 68, 0.4)',
              border: '1px solid rgba(239, 68, 68, 0.8)',
              fontSize: 12,
              fontWeight: 'bold',
              color: '#fef2f2'
            }}>!</span>
            出勤チェックが付いていないスタッフ {unmarkedStaffDetails.length}名
          </div>
          <div style={{
            fontSize: 11,
            color: 'rgba(252, 165, 165, 0.55)',
            lineHeight: 1.5,
            marginBottom: 8,
            paddingLeft: 28
          }}>
            {selectedDate !== 'all' ? `${selectedDate} に` : '選択期間内に'}サービス予定があるのに、出勤者チェックが付いていません。出勤者の入れ忘れ?
          </div>
          <div style={{
            display: 'flex',
            flexWrap: 'wrap',
            gap: 6,
            paddingLeft: 28
          }}>
            {unmarkedStaffDetails.map((d, i) => (
              <div key={i} style={{
                padding: '5px 12px',
                background: 'rgba(0, 0, 0, 0.3)',
                border: '1px solid rgba(239, 68, 68, 0.5)',
                borderRadius: 3,
                fontSize: 12,
                color: '#fef2f2'
              }}>
                <strong style={{ marginRight: 6 }}>{d.staff}</strong>
                <span style={{ color: 'rgba(252, 165, 165, 0.7)', fontSize: 11 }}>
                  {d.recordCount}件 / {d.userCount}名担当
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ===== アラート②: サービスがあるのに表に出てない利用者 ===== */}
      {missingDetails && missingDetails.length > 0 && (
        <div style={{
          marginBottom: 16,
          padding: '14px 18px',
          background: 'rgba(251, 146, 60, 0.1)',
          border: '1px solid rgba(251, 146, 60, 0.5)',
          borderRadius: 4,
          boxShadow: '0 0 20px rgba(251, 146, 60, 0.2), inset 0 0 10px rgba(251, 146, 60, 0.1)'
        }}>
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            marginBottom: 10,
            fontSize: 13,
            fontWeight: 600,
            color: '#fdba74'
          }}>
            <span style={{
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: 20,
              height: 20,
              borderRadius: '50%',
              background: 'rgba(251, 146, 60, 0.4)',
              border: '1px solid rgba(251, 146, 60, 0.8)',
              fontSize: 12,
              fontWeight: 'bold',
              color: '#fffbeb'
            }}>!</span>
            サービスがあるのに表に出てない利用者 {missingDetails.length}名
          </div>
          <div style={{
            fontSize: 11,
            color: 'rgba(253, 186, 116, 0.55)',
            lineHeight: 1.5,
            marginBottom: 8,
            paddingLeft: 28
          }}>
            この利用者のサービス担当が出勤者の中にいません。
          </div>
          <div style={{
            fontSize: 12,
            color: 'rgba(253, 186, 116, 0.95)',
            lineHeight: 1.8,
            paddingLeft: 28
          }}>
            {missingDetails.map((m, i) => (
              <div key={i} style={{
                marginTop: 4,
                padding: '4px 10px',
                background: 'rgba(0, 0, 0, 0.25)',
                borderLeft: '2px solid rgba(251, 146, 60, 0.7)',
                borderRadius: 2
              }}>
                <strong style={{
                  color: '#fffbeb',
                  marginRight: 8,
                  fontSize: 13
                }}>{m.user}様</strong>
                {m.otherStaff.length > 0 && (
                  <span style={{ fontSize: 11 }}>
                    担当:
                    <span style={{
                      color: '#fef3c7',
                      marginLeft: 6,
                      fontWeight: 500
                    }}>{m.otherStaff.join(' / ')}</span>
                  </span>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* テーブル */}
      <GlowFrame>
        <div style={{
          overflow: 'auto',
          maxHeight: 'calc(100vh - 280px)',
          padding: 0
        }}>
          <table style={{
            borderCollapse: 'collapse',
            fontSize: 10,
            whiteSpace: 'nowrap',
            color: '#e5e7eb'
          }}>
            <thead>
              <tr>
                <th style={{
                  position: 'sticky',
                  top: 0,
                  left: 0,
                  zIndex: 4,
                  background: '#0a0d12',
                  color: '#fef3c7',
                  padding: '8px 12px',
                  fontSize: 11,
                  fontWeight: 500,
                  letterSpacing: '0.1em',
                  minWidth: 100,
                  borderRight: '1px solid rgba(251, 191, 36, 0.3)',
                  borderBottom: '1px solid rgba(251, 191, 36, 0.3)'
                }}>
                  スタッフ / 時刻
                </th>
                {Array.from({ length: 24 }, (_, h) => (
                  Array.from({ length: 4 }, (_, q) => {
                    const isHourMark = q === 0;
                    return (
                      <th
                        key={`${h}-${q}`}
                        style={{
                          position: 'sticky',
                          top: 0,
                          zIndex: 3,
                          background: isHourMark ? 'rgba(251, 191, 36, 0.08)' : '#0a0d12',
                          color: isHourMark ? '#fef3c7' : 'transparent',
                          fontSize: 9,
                          fontWeight: 500,
                          padding: '4px 0',
                          minWidth: 22,
                          height: 22,
                          textAlign: 'center',
                          borderLeft: isHourMark ? '1px solid rgba(251, 191, 36, 0.3)' : 'none',
                          borderBottom: '1px solid rgba(251, 191, 36, 0.15)'
                        }}
                      >
                        {isHourMark ? h : ''}
                      </th>
                    );
                  })
                ))}
              </tr>
            </thead>
            <tbody>
              {staffList.map((staff, rowIdx) => {
                const shift = shiftSlotMap[staff];
                const hasShift = shift !== null;
                const staffType = staffTypes ? staffTypes[staff] : '';

                // コード行のレンダリング用に、連続する同groupIdスロットをまとめる
                // groupSegments: [{startCol, endCol, code, serviceClass, isOnShiftMostly}]
                const codeSegments = [];
                let i = 0;
                while (i < 96) {
                  const cell = grid[staff][i];
                  if (cell && cell.code) {
                    let j = i;
                    while (j < 96 && grid[staff][j] && grid[staff][j].groupId === cell.groupId) j++;
                    codeSegments.push({
                      startCol: i,
                      endCol: j - 1,
                      span: j - i,
                      code: cell.code,
                      serviceClass: cell.serviceClass
                    });
                    i = j;
                  } else {
                    i++;
                  }
                }

                return (
                  <React.Fragment key={staff}>
                    {/* 上行: 氏名 + 利用者名 */}
                    <tr>
                      <th
                        style={{
                          position: 'sticky',
                          left: 0,
                          zIndex: 2,
                          background: '#0a0d12',
                          color: '#fef3c7',
                          padding: '0 12px',
                          fontSize: 13,
                          fontWeight: 600,
                          textAlign: 'left',
                          borderRight: '1px solid rgba(251, 191, 36, 0.3)',
                          borderBottom: '1px dashed rgba(251, 191, 36, 0.15)',
                          verticalAlign: 'middle',
                          height: 28
                        }}
                      >
                        {staff}
                      </th>
                      {Array.from({ length: 96 }, (_, i) => {
                        const cell = grid[staff][i];
                        const isHourStart = i % 4 === 0;
                        const isHovered = hoveredCell && hoveredCell.row === rowIdx && hoveredCell.col === i;
                        const cellColor = cell ? SERVICE_COLORS[cell.serviceClass] : null;
                        const isOnShift = shift && i >= shift.startSlot && i < shift.endSlot;
                        let bgColor = 'transparent';
                        let shadowStyle = 'none';
                        if (cell) {
                          const baseAlpha = isOnShift ? 0.28 : 0.18;
                          const hoverAlpha = isOnShift ? 0.5 : 0.35;
                          bgColor = `rgba(${hexToRgb(cellColor)}, ${isHovered ? hoverAlpha : baseAlpha})`;
                          shadowStyle = `inset 0 0 ${isHovered ? 16 : 8}px rgba(${hexToRgb(cellColor)}, ${isHovered ? 0.5 : 0.25})`;
                        } else if (isOnShift) {
                          bgColor = `rgba(0, 255, 0, ${isHovered ? 0.18 : 0.1})`;
                        } else if (hasShift) {
                          bgColor = `rgba(255, 255, 153, ${isHovered ? 0.12 : 0.06})`;
                        }
                        return (
                          <td
                            key={i}
                            onMouseEnter={() => cell && setHoveredCell({ row: rowIdx, col: i, info: cell })}
                            onMouseLeave={() => setHoveredCell(null)}
                            style={{
                              minWidth: 22,
                              height: 28,
                              padding: cell ? '2px 3px' : 0,
                              textAlign: 'center',
                              background: bgColor,
                              color: cell ? '#fef3c7' : 'transparent',
                              fontSize: 9,
                              fontWeight: 500,
                              borderLeft: isHourStart
                                ? '1px solid rgba(251, 191, 36, 0.2)'
                                : '1px solid rgba(251, 191, 36, 0.04)',
                              borderBottom: '1px dashed rgba(251, 191, 36, 0.06)',
                              boxShadow: shadowStyle,
                              cursor: cell ? 'pointer' : 'default',
                              transition: 'all 0.2s',
                              overflow: 'hidden'
                            }}
                            title={cell ? `${staff} → ${cell.user} (${cell.service})` : ''}
                          >
                            {cell && cell.isStart ? cell.user : ''}
                          </td>
                        );
                      })}
                    </tr>
                    {/* 下行: 区分(社員/登録) + コード(連続スロットはcolspan結合) */}
                    <tr>
                      <th
                        style={{
                          position: 'sticky',
                          left: 0,
                          zIndex: 2,
                          background: '#0a0d12',
                          color: 'rgba(251, 191, 36, 0.75)',
                          padding: '0 12px',
                          fontSize: 11,
                          fontWeight: 400,
                          textAlign: 'left',
                          letterSpacing: '0.15em',
                          borderRight: '1px solid rgba(251, 191, 36, 0.3)',
                          borderBottom: '2px solid rgba(251, 191, 36, 0.2)',
                          verticalAlign: 'middle',
                          height: 24
                        }}
                      >
                        {staffType || '社員'}
                      </th>
                      {(() => {
                        const tds = [];
                        let col = 0;
                        let segIdx = 0;
                        while (col < 96) {
                          const seg = codeSegments[segIdx];
                          if (seg && seg.startCol === col) {
                            // 結合セル(サービスコード)
                            const segColor = SERVICE_COLORS[seg.serviceClass];
                            const segIsOnShift = shift && col >= shift.startSlot && col < shift.endSlot;
                            const baseAlpha = segIsOnShift ? 0.35 : 0.22;
                            tds.push(
                              <td
                                key={col}
                                colSpan={seg.span}
                                style={{
                                  minWidth: 22 * seg.span,
                                  height: 24,
                                  padding: '2px 3px',
                                  textAlign: 'center',
                                  background: `rgba(${hexToRgb(segColor)}, ${baseAlpha})`,
                                  color: '#fef3c7',
                                  fontSize: 10,
                                  fontWeight: 600,
                                  letterSpacing: '0.05em',
                                  borderLeft: col % 4 === 0
                                    ? '1px solid rgba(251, 191, 36, 0.2)'
                                    : '1px solid rgba(251, 191, 36, 0.04)',
                                  borderBottom: '2px solid rgba(251, 191, 36, 0.2)',
                                  boxShadow: `inset 0 0 6px rgba(${hexToRgb(segColor)}, 0.25)`,
                                  overflow: 'hidden'
                                }}
                              >
                                {seg.code}
                              </td>
                            );
                            col += seg.span;
                            segIdx++;
                          } else {
                            // コードなしセル: 勤務時間内/外の色のみ
                            const isOnShift = shift && col >= shift.startSlot && col < shift.endSlot;
                            const isHourStart = col % 4 === 0;
                            let bgColor = 'transparent';
                            if (isOnShift) {
                              bgColor = 'rgba(0, 255, 0, 0.1)';
                            } else if (hasShift) {
                              bgColor = 'rgba(255, 255, 153, 0.06)';
                            }
                            tds.push(
                              <td
                                key={col}
                                style={{
                                  minWidth: 22,
                                  height: 24,
                                  background: bgColor,
                                  borderLeft: isHourStart
                                    ? '1px solid rgba(251, 191, 36, 0.2)'
                                    : '1px solid rgba(251, 191, 36, 0.04)',
                                  borderBottom: '2px solid rgba(251, 191, 36, 0.2)'
                                }}
                              />
                            );
                            col++;
                          }
                        }
                        return tds;
                      })()}
                    </tr>
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </GlowFrame>

      {/* 凡例 */}
      <div style={{
        display: 'flex',
        gap: 16,
        marginTop: 16,
        padding: '10px 16px',
        background: 'rgba(20, 23, 33, 0.4)',
        border: '1px solid rgba(251, 191, 36, 0.1)',
        borderRadius: 4,
        fontSize: 11,
        flexWrap: 'wrap',
        alignItems: 'center'
      }}>
        <span style={{
          color: 'rgba(251, 191, 36, 0.5)',
          letterSpacing: '0.15em',
          fontSize: 10
        }}>SERVICE TYPE</span>
        {Object.entries(SERVICE_COLORS).map(([k, c]) => (
          <span key={k} style={{ display: 'flex', alignItems: 'center', gap: 6, color: 'rgba(229, 231, 235, 0.7)' }}>
            <span style={{
              width: 12, height: 12,
              background: `rgba(${hexToRgb(c)}, 0.3)`,
              boxShadow: `inset 0 0 6px rgba(${hexToRgb(c)}, 0.6), 0 0 6px rgba(${hexToRgb(c)}, 0.3)`,
              border: `1px solid rgba(${hexToRgb(c)}, 0.5)`,
              borderRadius: 2
            }} />
            {k}
          </span>
        ))}
        <span style={{
          marginLeft: 8,
          paddingLeft: 12,
          borderLeft: '1px solid rgba(251, 191, 36, 0.2)',
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          color: 'rgba(229, 231, 235, 0.7)'
        }}>
          <span style={{
            width: 12, height: 12,
            background: 'rgba(0, 255, 0, 0.15)',
            border: '1px solid rgba(0, 255, 0, 0.4)',
            borderRadius: 2
          }} />
          勤務時間内
        </span>
        <span style={{
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          color: 'rgba(229, 231, 235, 0.7)'
        }}>
          <span style={{
            width: 12, height: 12,
            background: 'rgba(255, 255, 153, 0.15)',
            border: '1px solid rgba(255, 255, 153, 0.4)',
            borderRadius: 2
          }} />
          勤務時間外
        </span>
      </div>
    </div>
  );
}

function SummaryItem({ label, value }) {
  return (
    <span style={{ display: 'flex', alignItems: 'baseline', gap: 6 }}>
      <span style={{ fontSize: 10, opacity: 0.7 }}>{label}</span>
      <strong style={{ color: '#fef3c7', fontSize: 14, fontWeight: 500 }}>{value}</strong>
    </span>
  );
}

function hexToRgb(hex) {
  const m = hex.replace('#', '').match(/.{2}/g);
  return m.map(h => parseInt(h, 16)).join(', ');
}

// ============ PDF 読取結果デバッグパネル ============
function DebugPanel({ bundles }) {
  const [fileIdx, setFileIdx] = useState(0);
  const bundle = bundles[fileIdx];
  if (!bundle) return null;

  return (
    <GlowFrame style={{ minHeight: 400 }}>
      <div style={{ padding: 18 }}>
        <div style={{
          fontSize: 11,
          color: '#fef3c7',
          letterSpacing: '0.2em',
          textTransform: 'uppercase',
          marginBottom: 14,
          paddingBottom: 10,
          borderBottom: '1px solid rgba(251, 191, 36, 0.15)'
        }}>
          PDF 読取結果
        </div>

        {/* ファイル切替 */}
        <div style={{
          display: 'flex',
          gap: 6,
          flexWrap: 'wrap',
          marginBottom: 12
        }}>
          {bundles.map((b, i) => (
            <button
              key={i}
              onClick={() => setFileIdx(i)}
              style={{
                padding: '4px 10px',
                fontSize: 10,
                background: i === fileIdx ? 'rgba(251, 191, 36, 0.2)' : 'rgba(0,0,0,0.3)',
                border: `1px solid rgba(251, 191, 36, ${i === fileIdx ? 0.5 : 0.2})`,
                color: '#fef3c7',
                borderRadius: 3,
                cursor: 'pointer',
                fontFamily: 'inherit'
              }}
            >
              {b.file} ({b.rowCount})
            </button>
          ))}
        </div>

        {/* メタ情報 */}
        <div style={{
          padding: 10,
          background: 'rgba(0,0,0,0.3)',
          border: '1px solid rgba(251, 191, 36, 0.15)',
          borderRadius: 3,
          marginBottom: 12,
          fontSize: 11,
          color: 'rgba(251, 191, 36, 0.7)',
          lineHeight: 1.6
        }}>
          <div>ファイル: <span style={{ color: '#fef3c7' }}>{bundle.file}</span></div>
          <div>検出スタッフ名: <span style={{ color: '#fef3c7' }}>{bundle.staff || '(未検出)'}</span></div>
          <div>抽出行数: <span style={{ color: '#fef3c7' }}>{bundle.rowCount}</span></div>
          {bundle.debug.notes.length > 0 && (
            <div style={{ marginTop: 6 }}>
              <div style={{ fontSize: 10, opacity: 0.7 }}>NOTES:</div>
              {bundle.debug.notes.slice(0, 5).map((n, i) => (
                <div key={i} style={{ fontSize: 10, marginLeft: 8 }}>· {n}</div>
              ))}
            </div>
          )}
        </div>

        {/* 行リスト */}
        <div style={{
          fontSize: 10,
          color: 'rgba(251, 191, 36, 0.5)',
          marginBottom: 6,
          letterSpacing: '0.1em'
        }}>
          PDFから読み取れた行 ({bundle.debug.lines.length}行)
        </div>
        <div style={{
          maxHeight: 'calc(100vh - 430px)',
          overflowY: 'auto',
          background: 'rgba(0,0,0,0.4)',
          border: '1px solid rgba(251, 191, 36, 0.1)',
          borderRadius: 3,
          padding: 10
        }}>
          {bundle.debug.lines.map((line, i) => (
            <div
              key={i}
              style={{
                display: 'grid',
                gridTemplateColumns: '40px 40px 1fr',
                gap: 8,
                padding: '3px 0',
                fontSize: 11,
                fontFamily: 'monospace',
                color: 'rgba(229, 231, 235, 0.7)',
                borderBottom: '1px solid rgba(251, 191, 36, 0.05)'
              }}
            >
              <span style={{ color: 'rgba(251, 191, 36, 0.4)' }}>p{line.page}</span>
              <span style={{ color: 'rgba(251, 191, 36, 0.4)' }}>y{line.y}</span>
              <span style={{ wordBreak: 'break-all' }}>{line.text}</span>
            </div>
          ))}
        </div>

        <div style={{
          marginTop: 14,
          padding: 10,
          background: 'rgba(251, 191, 36, 0.05)',
          border: '1px solid rgba(251, 191, 36, 0.2)',
          borderRadius: 3,
          fontSize: 11,
          color: '#fde68a',
          lineHeight: 1.6
        }}>
          上記の「PDFから読み取れた行」を確認して、<br/>
          ・スタッフ名がどう書かれているか<br/>
          ・利用者名・時間・サービス種別が同じ行に並んでいるか<br/>
          を教えてください。それに合わせて抽出ロジックを調整します。
        </div>
      </div>
    </GlowFrame>
  );
}
