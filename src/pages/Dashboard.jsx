import React, { useState } from 'react';

// ============ ダッシュボード ============
// アプリのトップ画面。機能ごとのカードを並べ、押すと各機能の画面に遷移する。
// 将来機能が増えたら CARDS 配列に追加するだけ。

export default function Dashboard({ onNavigate }) {
  const cards = [
    {
      id: 'schedule',
      icon: '📋',
      title: 'スタッフ・利用者 対応表',
      description: 'PDF/CSVを読み込み、シフト表をひな形保持でExcel出力',
    },
    {
      id: 'register',
      icon: '👥',
      title: 'スタッフ登録',
      description: 'スタッフの氏名・区分(社員/登録)を保存。シフト出力で自動補完',
    },
  ];

  return (
    <div
      style={{
        minHeight: '100vh',
        padding: '60px 40px',
        background: 'linear-gradient(180deg, #0a0d12 0%, #14181f 100%)',
        color: '#fef3c7',
        fontFamily: '"Yu Gothic", "Hiragino Sans", sans-serif',
      }}
    >
      <header style={{ marginBottom: 48, textAlign: 'center' }}>
        <h1
          style={{
            fontSize: 32,
            fontWeight: 300,
            letterSpacing: '0.2em',
            margin: 0,
            color: '#fef3c7',
            textShadow: '0 0 24px rgba(251, 191, 36, 0.4)',
          }}
        >
          シフト出力ツール
        </h1>
        <p
          style={{
            fontSize: 12,
            color: 'rgba(251, 191, 36, 0.5)',
            letterSpacing: '0.4em',
            marginTop: 8,
            textTransform: 'uppercase',
          }}
        >
          Staff Schedule Dashboard
        </p>
      </header>

      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
          gap: 24,
          maxWidth: 880,
          margin: '0 auto',
        }}
      >
        {cards.map((c) => (
          <DashboardCard
            key={c.id}
            icon={c.icon}
            title={c.title}
            description={c.description}
            onClick={() => onNavigate(c.id)}
          />
        ))}
      </div>
    </div>
  );
}

function DashboardCard({ icon, title, description, onClick }) {
  const [hover, setHover] = useState(false);
  return (
    <button
      onClick={onClick}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        textAlign: 'left',
        padding: '28px 24px',
        background: hover ? 'rgba(251, 191, 36, 0.1)' : 'rgba(0,0,0,0.4)',
        border: `1px solid rgba(251, 191, 36, ${hover ? 0.6 : 0.25})`,
        borderRadius: 6,
        color: '#fef3c7',
        fontFamily: 'inherit',
        cursor: 'pointer',
        transition: 'all 0.25s',
        boxShadow: hover ? '0 0 24px rgba(251, 191, 36, 0.3)' : 'none',
      }}
    >
      <div style={{ fontSize: 36, marginBottom: 12 }}>{icon}</div>
      <div
        style={{
          fontSize: 16,
          fontWeight: 500,
          letterSpacing: '0.1em',
          marginBottom: 8,
        }}
      >
        {title}
      </div>
      <div
        style={{
          fontSize: 11,
          color: 'rgba(251, 191, 36, 0.55)',
          letterSpacing: '0.05em',
          lineHeight: 1.6,
        }}
      >
        {description}
      </div>
    </button>
  );
}
