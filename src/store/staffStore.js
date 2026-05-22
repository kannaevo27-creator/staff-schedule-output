// ============ スタッフ情報ストア (localStorage) ============
// 登録済みスタッフを localStorage に保存・取得する。
// 同じブラウザで開けば残り続けるので、毎日 区分(社員/登録) を入れ直す手間が省ける。
// PCを変える/別ブラウザで開くと共有されない (=各PCで個別管理)

const STORAGE_KEY = 'registeredStaff';

// スタッフ1人 = { name: string, type: '社員' | '登録' }

export function loadStaff() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function saveStaff(list) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
}

export function addStaff(name, type) {
  const trimmed = String(name || '').trim();
  if (!trimmed) return loadStaff();
  const list = loadStaff();
  // 重複登録は最新で上書き (区分の変更にも対応)
  const idx = list.findIndex((s) => s.name === trimmed);
  if (idx >= 0) {
    list[idx] = { name: trimmed, type };
  } else {
    list.push({ name: trimmed, type });
  }
  saveStaff(list);
  return list;
}

export function removeStaff(name) {
  const list = loadStaff().filter((s) => s.name !== name);
  saveStaff(list);
  return list;
}

// 検出スタッフ名 → 登録区分を引く ('社員' | '登録' | null)
export function lookupType(name) {
  const list = loadStaff();
  const found = list.find((s) => s.name === name);
  return found ? found.type : null;
}
