import React, { useMemo, useState } from 'react';
import { useApp } from '../store/AppContext';
import { ScheduleStatus, ScheduleItem } from '../types';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import *as XLSX from 'xlsx';
import { Download, AlertCircle, X } from 'lucide-react';
import { format } from 'date-fns';
import { parseLocal } from '../utils';

const Statistics: React.FC = () => {
  const { teachers, schedules, subjects, classes } = useApp();
  const [showAlert, setShowAlert] = useState(true);

  // 1. Missed classes needing makeup
  // Logic: Filter out if Subject is Finished OR if there are corresponding Makeup sessions
  const missedClasses = useMemo(() => {
    const missedMap: Record<string, ScheduleItem[]> = {};
    const makeupMap: Record<string, number> = {};
    const completionCache: Record<string, boolean> = {};

    // Helper to check if subject is completed
    const isSubjectFinished = (subId: string, clsId: string) => {
        const key = `${subId}-${clsId}`;
        if (completionCache[key] !== undefined) return completionCache[key];

        const subject = subjects.find(s => s.id === subId);
        if (!subject) return false;

        const learned = schedules.filter(s => 
            s.subjectId === subId && 
            s.classId === clsId && 
            s.status !== ScheduleStatus.OFF
        ).reduce((acc, curr) => acc + curr.periodCount, 0);

        const finished = learned >= subject.totalPeriods;
        completionCache[key] = finished;
        return finished;
    };

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
        
        // 1. If subject completed -> No alert
        if (isSubjectFinished(subId, clsId)) return;

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

  // 2. Teacher Stats
  const teacherStats = teachers.map(t => {
    // Logic: Count classes OR (Exams marked as 'thực hành')
    const taught = schedules.filter(s => 
        s.teacherId === t.id && 
        (s.status === ScheduleStatus.COMPLETED || s.status === ScheduleStatus.ONGOING || s.status === ScheduleStatus.PENDING) &&
        (s.type === 'class' || (s.type === 'exam' && s.note?.toLowerCase().includes('thực hành')))
    ); 
    const totalPeriods = taught.reduce((acc, curr) => acc + curr.periodCount, 0);
    return {
      name: t.name,
      periods: totalPeriods,
      income: totalPeriods * t.ratePerPeriod
    };
  }).filter(t => t.periods > 0); 

  // 3. Subject Progress (All active subjects across all classes)
  const subjectStats = useMemo(() => {
    const results: any[] = [];
    
    classes.forEach(cls => {
        // Subjects for this class based on major
        const classSubjects = subjects.filter(s => s.majorId === cls.majorId);
        
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
     const data = schedules
        .filter(s => s.status === ScheduleStatus.COMPLETED || s.status === ScheduleStatus.ONGOING)
        .map(s => ({
            'Ngày dạy': format(parseLocal(s.date), 'dd/MM/yyyy'),
            'Giáo viên': teachers.find(t => t.id === s.teacherId)?.name,
            'Môn học': subjects.find(sub => sub.id === s.subjectId)?.name,
            'Loại': s.type === 'exam' ? 'Thi' : 'Học',
            'Ghi chú': s.note || '',
            'Số tiết': s.periodCount,
            'Lớp': classes.find(c => c.id === s.classId)?.name || s.classId
        }));
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ThongKeGiangDay");
    XLSX.writeFile(wb, "Thong_Ke_Giang_Vien.xlsx");
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
              <h3 className="font-bold text-gray-700">Số tiết dạy theo Giáo viên</h3>
              <button onClick={exportTeacherReport} className="text-sm bg-green-100 text-green-700 px-2 py-1 rounded flex items-center hover:bg-green-200">
                <Download size={14} className="mr-1"/> Xuất Excel
              </button>
           </div>
           <p className="text-xs text-gray-500 mb-2 italic">* Chỉ tính các tiết học và các buổi thi được đánh dấu là "thực hành".</p>
           {teacherStats.length > 0 ? (
             <div className="flex-1 min-h-0">
               <ResponsiveContainer width="100%" height="100%">
                 <BarChart data={teacherStats}>
                   <CartesianGrid strokeDasharray="3 3" vertical={false} />
                   <XAxis dataKey="name" tick={{fontSize: 11}} interval={0} angle={-15} textAnchor="end" height={60} />
                   <YAxis />
                   <Tooltip />
                   <Bar dataKey="periods" fill="#3B82F6" name="Số tiết" barSize={40} />
                 </BarChart>
               </ResponsiveContainer>
             </div>
           ) : (
             <div className="flex-1 flex items-center justify-center text-gray-400 italic">
               Chưa có dữ liệu giáo viên đang dạy.
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
