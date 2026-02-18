
import React, { useMemo, useState } from 'react';
import { useApp } from '../store/AppContext';
import { ScheduleStatus, ScheduleItem } from '../types';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import XLSX from 'xlsx';
import { Download, AlertCircle, X } from 'lucide-react';
import { format } from 'date-fns';
import { parseLocal, isSubjectFinished } from '../utils';

const Statistics: React.FC = () => {
  const { teachers, schedules, subjects, classes } = useApp();
  const [showAlert, setShowAlert] = useState(true);

  // 1. Missed classes needing makeup
  // Logic: Filter out if Subject is Finished OR if there are corresponding Makeup sessions
  const missedClasses = useMemo(() => {
    const missedMap: Record<string, ScheduleItem[]> = {};
    const makeupMap: Record<string, number> = {};
    
    // Group items
    schedules.forEach(s => {
        const key = `${s.subjectId}-${s.classId}`;
        if (s.status === ScheduleStatus.OFF) {
            if (!missedMap[key]) missedMap[key] = [];
            missedMap[key].push(s);
        }
        if (s.status === ScheduleStatus.MAKEUP) {
            makeupMap[key] = (makeupMap[key] || 0) + 1;
        }
    });

    let results: ScheduleItem[] = [];

    Object.keys(missedMap).forEach(key => {
        const [subId, clsId] = key.split('-');
        
        const subject = subjects.find(s => s.id === subId);
        if (!subject) return;

        // 1. If subject completed (use shared helper) -> No alert
        if (isSubjectFinished(subject, clsId, schedules)) return;

        // 2. If has makeup sessions -> Reduce alerts
        // We assume makeup sessions cover the earliest missed classes first
        const makeupsCount = makeupMap[key] || 0;
        const missedItems = missedMap[key].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
        
        // If we have enough makeups to cover all missed, show nothing.
        // Otherwise show the remaining earliest ones (or latest? usually earliest are the ones needing attention).
        if (makeupsCount >= missedItems.length) return;

        const remaining = missedItems.slice(makeupsCount);
        results = [...results, ...remaining];
    });

    return results;
  }, [schedules, subjects]);

  // 2. Teacher Stats (Memoized)
  const teacherStats = useMemo(() => {
      return teachers.map(t => {
        // Base filter: Get all valid teaching schedules for this teacher
        // Exclude OFF sessions
        // Include 'class' type OR 'exam' type if it's 'thực hành'
        const allSchedules = schedules.filter(s => 
            s.teacherId === t.id && 
            s.status !== ScheduleStatus.OFF &&
            (s.type === 'class' || (s.type === 'exam' && s.note?.toLowerCase().includes('thực hành')))
        ); 

        // Metric 1: Active Load (For Chart)
        // Sum of periods for subjects that are NOT YET FINISHED.
        const activeLoad = allSchedules.reduce((acc, s) => {
            const sub = subjects.find(sub => sub.id === s.subjectId);
            // If subject exists and is NOT finished, count it towards active load
            if (sub && !isSubjectFinished(sub, s.classId, schedules)) {
                return acc + s.periodCount;
            }
            return acc;
        }, 0);

        // Metric 2: Taught Load (For Excel / History)
        // Sum of periods that are explicitly marked as COMPLETED (actually taught).
        const taughtLoad = allSchedules.reduce((acc, s) => {
            if (s.status === ScheduleStatus.COMPLETED) {
                return acc + s.periodCount;
            }
            return acc;
        }, 0);

        return {
          name: t.name,
          activePeriods: activeLoad,
          taughtPeriods: taughtLoad
        };
      });
  }, [teachers, schedules, subjects]);

  // Filter for Chart: Only show teachers currently teaching active subjects
  const chartData = teacherStats.filter(t => t.activePeriods > 0);

  // 3. Subject Progress (All active subjects across all classes)
  const subjectStats = useMemo(() => {
    const results: any[] = [];
    
    classes.forEach(cls => {
        const isH8 = cls.name.toUpperCase().includes('H8');
        
        // Subjects for this class: Major specific OR Common OR Culture (if not H8) OR Culture 8 (if H8)
        const classSubjects = subjects.filter(s => {
            if (s.majorId === 'common') return true;
            if (s.majorId === 'culture') return !isH8;
            if (s.majorId === 'culture_8') return isH8;
            return s.majorId === cls.majorId;
        });
        
        classSubjects.forEach(sub => {
             const relevantSchedules = schedules.filter(sch => 
                sch.subjectId === sub.id && 
                sch.classId === cls.id && 
                sch.status !== ScheduleStatus.OFF
            );
            
            const learned = relevantSchedules.reduce((acc, curr) => acc + curr.periodCount, 0);
            const remaining = Math.max(0, sub.totalPeriods - learned);
            
            // Filter: Only show subjects that are "currently learning"
            // Condition: Learned > 0 (started) AND Remaining > 0 (not finished)
            if (learned > 0 && remaining > 0) {
                 results.push({
                     name: sub.name,
                     className: cls.name,
                     // Combine for unique label
                     fullName: `${sub.name} (${cls.name})`,
                     total: sub.totalPeriods,
                     learned: learned,
                     remaining: remaining
                 });
            }
        });
    });
    
    return results;
  }, [subjects, schedules, classes]);

  const exportTeacherReport = () => {
     // Prepare data for all teachers (even those with 0 active periods)
     const data = teacherStats.map(t => ({
         'Họ và tên': t.name,
         'Số tiết đang dạy (Môn chưa kết thúc)': t.activePeriods,
         'Số tiết đã dạy (Thực tế đã hoàn thành)': t.taughtPeriods,
     }));
    
    const ws = XLSX.utils.json_to_sheet(data);
    
    // Set column widths
    ws['!cols'] = [{ wch: 25 }, { wch: 35 }, { wch: 35 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ThongKeGiaoVien");
    XLSX.writeFile(wb, "Thong_Ke_Tiet_Day_Giao_Vien.xlsx");
  };

  // Dynamic height calculation
  const progressChartHeight = Math.max(300, subjectStats.length * 60);

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold text-gray-800">Thống kê báo cáo</h1>

      {/* Quick Alert: Missed Classes */}
      {missedClasses.length > 0 && showAlert && (
        <div className="bg-red-50 border border-red-200 rounded-lg p-4 flex items-start space-x-3 relative">
          <AlertCircle className="text-red-500 mt-1 flex-shrink-0" />
          <div className="flex-1">
             <h3 className="font-bold text-red-700">Cần xếp lịch bù ({missedClasses.length} buổi)</h3>
             <ul className="text-sm text-red-600 mt-1 list-disc pl-4">
               {missedClasses.map(m => (
                 <li key={m.id}>
                    {m.date} - GV {teachers.find(t => t.id === m.teacherId)?.name} - Môn {subjects.find(s => s.id === m.subjectId)?.name} (Lớp {classes.find(c => c.id === m.classId)?.name})
                 </li>
               ))}
             </ul>
          </div>
          <button 
            onClick={() => setShowAlert(false)} 
            className="text-red-400 hover:text-red-600 p-1 hover:bg-red-100 rounded absolute top-2 right-2"
            title="Đóng thông báo"
          >
            <X size={20} />
          </button>
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Teacher Chart */}
        <div className="bg-white p-6 rounded-xl shadow border h-[500px] flex flex-col">
           <div className="flex justify-between items-center mb-4 flex-shrink-0">
              <h3 className="font-bold text-gray-700">Số tiết đang dạy (Môn chưa kết thúc)</h3>
              <button onClick={exportTeacherReport} className="text-sm bg-green-100 text-green-700 px-3 py-1.5 rounded flex items-center hover:bg-green-200 transition-colors font-medium">
                <Download size={14} className="mr-1"/> Xuất Thống Kê Tổng
              </button>
           </div>
           <p className="text-xs text-gray-500 mb-2 italic">
               * Biểu đồ chỉ hiển thị tải công việc hiện tại (các môn chưa kết thúc). <br/>
               * Nhấn "Xuất Thống Kê Tổng" để xem cả số tiết đã dạy (lịch sử) của toàn bộ giáo viên.
           </p>
           {chartData.length > 0 ? (
             <div className="flex-1 min-h-0">
               <ResponsiveContainer width="100%" height="100%">
                 <BarChart data={chartData}>
                   <CartesianGrid strokeDasharray="3 3" vertical={false} />
                   <XAxis dataKey="name" tick={{fontSize: 11}} interval={0} angle={-15} textAnchor="end" height={60} />
                   <YAxis />
                   <Tooltip cursor={{fill: '#f3f4f6'}} />
                   <Bar dataKey="activePeriods" fill="#3B82F6" name="Số tiết đang dạy" barSize={40} />
                 </BarChart>
               </ResponsiveContainer>
             </div>
           ) : (
             <div className="flex-1 flex items-center justify-center text-gray-400 italic">
               Không có giáo viên nào đang dạy môn chưa kết thúc.
             </div>
           )}
        </div>

        {/* Progress Chart (Scrollable) */}
         <div className="bg-white p-6 rounded-xl shadow border h-[500px] flex flex-col">
           <h3 className="font-bold text-gray-700 mb-4 flex-shrink-0">Tiến độ (Tất cả các môn đang học)</h3>
           
           <div className="flex-1 overflow-y-auto pr-2 custom-scrollbar">
               {subjectStats.length > 0 ? (
                  <div style={{ height: progressChartHeight }}>
                    <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={subjectStats} layout="vertical" margin={{ top: 5, right: 30, left: 10, bottom: 5 }}>
                            <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                            <XAxis type="number" />
                            <YAxis dataKey="fullName" type="category" width={180} tick={{fontSize: 11}} />
                            <Tooltip cursor={{fill: '#f3f4f6'}} />
                            <Legend verticalAlign="top" height={36} />
                            <Bar dataKey="learned" stackId="a" fill="#10B981" name="Đã học" barSize={24} />
                            <Bar dataKey="remaining" stackId="a" fill="#E5E7EB" name="Chưa học" barSize={24} radius={[0, 4, 4, 0]} />
                        </BarChart>
                    </ResponsiveContainer>
                  </div>
               ) : (
                   <div className="h-full flex items-center justify-center text-gray-400 italic px-6 text-center">
                       Không có môn nào đang diễn ra (Tất cả đã xong hoặc chưa bắt đầu).
                   </div>
               )}
           </div>
        </div>
      </div>
    </div>
  );
};

export default Statistics;
