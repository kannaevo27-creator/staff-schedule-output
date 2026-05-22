import React, { useState } from 'react';
import Dashboard from './pages/Dashboard';
import StaffSchedule from './pages/StaffSchedule';
import StaffRegister from './pages/StaffRegister';

// ============ アプリ全体のルーター ============
// URL ではなく React state で画面を切り替える最小構成。
// 機能を追加する場合は: 1) pages/ に画面コンポーネント追加, 2) Dashboard.jsx の cards に追加,
// 3) ここに分岐を追加 の3点で済む。

export default function App() {
  const [view, setView] = useState('dashboard'); // 'dashboard' | 'schedule' | 'register'

  if (view === 'schedule') {
    return <StaffSchedule onBack={() => setView('dashboard')} />;
  }
  if (view === 'register') {
    return <StaffRegister onBack={() => setView('dashboard')} />;
  }
  return <Dashboard onNavigate={setView} />;
}
