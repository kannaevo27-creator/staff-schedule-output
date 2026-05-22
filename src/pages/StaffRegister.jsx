import React, { useState } from 'react';
import { loadStaff, addStaff, removeStaff } from '../store/staffStore';

// ============ スタッフ登録画面 ============
// 氏名と区分(社員/登録)をブラウザに保存。
// シフト出力時、登録済スタッフは区分が自動補完される。

export default function StaffRegister({ onBack }) {
  const [list, setList] = useState(() => loadStaff());
  const [name, setName] = useState('');
  const [type, setType] = useState('社員');

  function handleAdd() {
    const t = name.trim();
    if (!t) return;
    const newList = addStaff(t, type);
    setList(newList);
    setName('');
  }

  function handleRemove(n) {
    if (!confirm(`「${n}」を削除しますか？`)) return;
    setList(removeStaff(n));
  }

  function handleKeyDown(e) {
    if (e.key === 'Enter') handleAdd();
  }

  return (
    <div
      style={{
        minHeight: '100vh',
        padding: '40px 32px',
        background: 'linear-gradient(180deg, #0a0d12 0%, #14181f 100%)',
        color: '#fef3c7',
        fontFamily: '"Yu Gothic", "Hiragino Sans", sans-serif',
      }}
    >
      <header style={{ marginBottom: 24, position: 'relative' }}>
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
            cursor: 'pointer',
          }}
        >← ダッシュボード</button>
        <h1
          style={{
            fontSize: 24,
            fontWeight: 300,
            letterSpacing: '0.15em',
            margin: 0,
            marginTop: 28,
            color: '#fef3c7',
            textShadow: '0 0 18px rgba(251, 191, 36, 0.35)',
          }}
        >
          スタッフ登録
        </h1>
        <p
          style={{
            fontSize: 11,
            color: 'rgba(251, 191, 36, 0.5)',
            letterSpacing: '0.3em',
            marginTop: 6,
            textTransform: 'uppercase',
          }}
        >
          Staff Master
        </p>
      </header>

      {/* 追加フォーム */}
      <section
        style={{
          maxWidth: 640,
          padding: 20,
          background: 'rgba(0,0,0,0.4)',
          border: '1px solid rgba(251, 191, 36, 0.25)',
          borderRadius: 5,
          marginBottom: 32,
        }}
      >
        <div
          style={{
            fontSize: 11,
            color: 'rgba(251, 191, 36, 0.6)',
            letterSpacing: '0.2em',
            marginBottom: 12,
          }}
        >
          新規追加
        </div>
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', flexWrap: 'wrap' }}>
          <input
            type="text"
            value={name}
            onChange={(e) => setName(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="スタッフ氏名"
            style={{
              flex: '1 1 200px',
              padding: '10px 12px',
              background: 'rgba(0,0,0,0.5)',
              border: '1px solid rgba(251, 191, 36, 0.35)',
              borderRadius: 3,
              color: '#fef3c7',
              fontFamily: 'inherit',
              fontSize: 14,
              letterSpacing: '0.05em',
              outline: 'none',
            }}
          />
          <div style={{ display: 'flex', gap: 8 }}>
            {['社員', '登録'].map((t) => (
              <label
                key={t}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: 4,
                  padding: '8px 12px',
                  background: type === t ? 'rgba(251, 191, 36, 0.2)' : 'rgba(0,0,0,0.3)',
                  border: `1px solid rgba(251, 191, 36, ${type === t ? 0.6 : 0.25})`,
                  borderRadius: 3,
                  cursor: 'pointer',
                  fontSize: 13,
                  letterSpacing: '0.1em',
                  userSelect: 'none',
                }}
              >
                <input
                  type="radio"
                  name="staff-type"
                  value={t}
                  checked={type === t}
                  onChange={() => setType(t)}
                  style={{ display: 'none' }}
                />
                {t}
              </label>
            ))}
          </div>
          <button
            onClick={handleAdd}
            disabled={!name.trim()}
            style={{
              padding: '10px 20px',
              background: name.trim() ? 'rgba(34, 197, 94, 0.2)' : 'rgba(100, 100, 100, 0.2)',
              border: `1px solid rgba(34, 197, 94, ${name.trim() ? 0.6 : 0.2})`,
              color: name.trim() ? '#bbf7d0' : 'rgba(229, 231, 235, 0.4)',
              fontFamily: 'inherit',
              fontSize: 13,
              letterSpacing: '0.15em',
              borderRadius: 3,
              cursor: name.trim() ? 'pointer' : 'default',
            }}
          >
            ＋ 追加
          </button>
        </div>
      </section>

      {/* 一覧 */}
      <section style={{ maxWidth: 640 }}>
        <div
          style={{
            fontSize: 11,
            color: 'rgba(251, 191, 36, 0.6)',
            letterSpacing: '0.2em',
            marginBottom: 12,
          }}
        >
          登録済スタッフ（{list.length}人）
        </div>
        {list.length === 0 ? (
          <div
            style={{
              padding: 40,
              textAlign: 'center',
              color: 'rgba(251, 191, 36, 0.35)',
              fontSize: 12,
              letterSpacing: '0.1em',
              border: '1px dashed rgba(251, 191, 36, 0.2)',
              borderRadius: 5,
            }}
          >
            まだスタッフが登録されていません
          </div>
        ) : (
          <div
            style={{
              background: 'rgba(0,0,0,0.4)',
              border: '1px solid rgba(251, 191, 36, 0.15)',
              borderRadius: 5,
              overflow: 'hidden',
            }}
          >
            {list.map((s, i) => (
              <div
                key={s.name}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  padding: '12px 16px',
                  borderTop: i > 0 ? '1px solid rgba(251, 191, 36, 0.1)' : 'none',
                  gap: 12,
                }}
              >
                <span style={{ flex: 1, fontSize: 14, letterSpacing: '0.05em' }}>{s.name}</span>
                <span
                  style={{
                    fontSize: 11,
                    padding: '3px 10px',
                    background:
                      s.type === '社員'
                        ? 'rgba(59, 130, 246, 0.18)'
                        : 'rgba(168, 85, 247, 0.18)',
                    border: `1px solid ${
                      s.type === '社員'
                        ? 'rgba(59, 130, 246, 0.5)'
                        : 'rgba(168, 85, 247, 0.5)'
                    }`,
                    color: s.type === '社員' ? '#93c5fd' : '#d8b4fe',
                    borderRadius: 3,
                    letterSpacing: '0.1em',
                  }}
                >
                  {s.type}
                </span>
                <button
                  onClick={() => handleRemove(s.name)}
                  style={{
                    padding: '4px 10px',
                    background: 'rgba(239, 68, 68, 0.12)',
                    border: '1px solid rgba(239, 68, 68, 0.4)',
                    color: '#fca5a5',
                    fontFamily: 'inherit',
                    fontSize: 11,
                    letterSpacing: '0.1em',
                    borderRadius: 3,
                    cursor: 'pointer',
                  }}
                >削除</button>
              </div>
            ))}
          </div>
        )}
      </section>
    </div>
  );
}
