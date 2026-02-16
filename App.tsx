
import React, { useState } from 'react';
import { AppProvider } from './store/AppContext';
import Dashboard from './components/Dashboard';
import ScheduleManager from './components/ScheduleManager';
import Statistics from './components/Statistics';
import Management from './components/Management';
import StudentManager from './components/StudentManager';
import SystemManager from './components/SystemManager';
import TeachingProgress from './components/TeachingProgress';
import Payment from './components/Payment';
import DocumentManager from './components/DocumentManager';
import HolidayManager from './components/HolidayManager';
import { LayoutDashboard, CalendarDays, PieChart, GraduationCap, Menu, X, Users, Settings, TrendingUp, CreditCard, FolderOpen, CalendarOff } from 'lucide-react';

const App: React.FC = () => {
  const [activeView, setActiveView] = useState<'dashboard' | 'schedule' | 'stats' | 'progress' | 'manage' | 'students' | 'system' | 'payment' | 'documents' | 'holidays'>('dashboard');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  const NavItem = ({ view, icon: Icon, label }: { view: typeof activeView, icon: any, label: string }) => (
    <button
      onClick={() => { setActiveView(view); setIsMobileMenuOpen(false); }}
      className={`flex items-center space-x-3 w-full p-3 rounded-lg transition-colors ${
        activeView === view ? 'bg-blue-600 text-white shadow-md' : 'text-gray-600 hover:bg-gray-100'
      }`}
    >
      <Icon size={20} />
      <span className="font-medium">{label}</span>
    </button>
  );

  return (
    <AppProvider>
      <div className="flex h-screen bg-gray-50 overflow-hidden font-sans">
        {/* Sidebar */}
        <aside className={`fixed inset-y-0 left-0 z-50 w-64 bg-white border-r transform transition-transform duration-200 ease-in-out md:translate-x-0 md:relative md:inset-auto md:flex md:flex-col ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full'}`}>
          <div className="h-16 flex items-center justify-center border-b">
            <h1 className="text-xl font-bold text-blue-800 flex items-center">
              <GraduationCap className="mr-2" /> EduPro
            </h1>
          </div>
          <nav className="flex-1 p-4 space-y-2 overflow-y-auto">
            <NavItem view="dashboard" icon={LayoutDashboard} label="Tổng quan" />
            <NavItem view="schedule" icon={CalendarDays} label="Quản lý lịch & Thi" />
            <NavItem view="holidays" icon={CalendarOff} label="Quản lý ngày nghỉ" />
            <NavItem view="payment" icon={CreditCard} label="Thanh toán giảng dạy" />
            <NavItem view="progress" icon={TrendingUp} label="Tiến độ giảng dạy" />
            <NavItem view="stats" icon={PieChart} label="Thống kê" />
            <NavItem view="manage" icon={GraduationCap} label="Quản lý giảng dạy" />
            <NavItem view="students" icon={Users} label="Quản lý HSSV" />
            <NavItem view="documents" icon={FolderOpen} label="Hồ sơ" />
            <NavItem view="system" icon={Settings} label="Hệ thống" />
          </nav>
          <div className="p-4 border-t text-xs text-center text-gray-400">
            v1.1.0
          </div>
        </aside>

        {/* Main Content */}
        <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
          {/* Mobile Header */}
          <header className="md:hidden bg-white h-16 border-b flex items-center justify-between px-4">
             <h1 className="font-bold text-blue-800">EduPro</h1>
             <div className="flex items-center gap-2">
                 <button onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}>
                   {isMobileMenuOpen ? <X /> : <Menu />}
                 </button>
             </div>
          </header>

          <main className="flex-1 overflow-auto p-4 md:p-8">
            <div className="max-w-7xl mx-auto">
              {activeView === 'dashboard' && <Dashboard />}
              {activeView === 'schedule' && <ScheduleManager />}
              {activeView === 'holidays' && <HolidayManager />}
              {activeView === 'progress' && <TeachingProgress />}
              {activeView === 'payment' && <Payment />}
              {activeView === 'documents' && <DocumentManager />}
              {activeView === 'stats' && <Statistics />}
              {activeView === 'manage' && <Management />}
              {activeView === 'students' && <StudentManager />}
              {activeView === 'system' && <SystemManager />}
            </div>
          </main>
        </div>
        
        {/* Mobile Overlay */}
        {isMobileMenuOpen && (
          <div 
            className="fixed inset-0 bg-black/50 z-40 md:hidden"
            onClick={() => setIsMobileMenuOpen(false)}
          />
        )}
      </div>
    </AppProvider>
  );
};

export default App;
