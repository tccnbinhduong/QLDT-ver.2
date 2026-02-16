import React, { useEffect, useState, useMemo } from 'react';
import { useApp } from '../store/AppContext';
import { BookOpen, Users, Calendar, Clock, AlertTriangle, CheckCircle } from 'lucide-react';
import { format, isSameDay } from 'date-fns';
import { vi } from 'date-fns/locale/vi';
import { parseLocal, determineStatus } from '../utils';
import { ScheduleStatus } from '../types';

const Dashboard: React.FC = () => {
  const { subjects, teachers, schedules, classes } = useApp();
  const [today] = useState(new Date());
  const [currentTime, setCurrentTime] = useState(new Date());

  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 60000);
    return () => clearInterval(timer);
  }, []);

  // Helper for date comparison (reset time to 00:00:00)
  const startOfToday = new Date(today);
  startOfToday.setHours(0, 0, 0, 0);

  // Filter today's schedules: Must be same day AND type 'class' (Exams excluded)
  const todaySchedules = schedules
    .filter(s => s.type === 'class' && isSameDay(parseLocal(s.date), today))
    .sort((a, b) => a.startPeriod - b.startPeriod);

  // Upcoming exams: type 'exam' AND date >= startOfToday (includes today)
  const upcomingExams = schedules
    .filter(s => s.type === 'exam' && parseLocal(s.date) >= startOfToday)
    .sort((a, b) => parseLocal(a.date).getTime() - parseLocal(b.date).getTime())
    .slice(0, 5);

  // Logic for completed subjects (Read-only, matches Statistics logic)
  // We read 'paid_completed_subjects' from localStorage to exclude deleted items
  const [paidItems] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('paid_completed_subjects');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const completedSubjects = useMemo(() => {
    const results: any[] = [];
    
    classes.forEach(cls => {
        const classSubjects = subjects.filter(s => s.majorId === cls.majorId);
        
        classSubjects.forEach(sub => {
            // Unique key to match with Statistics (deleted/hidden items)
            const uniqueKey = `${sub.id}-${cls.id}`;

            // Skip if hidden/deleted
            if (paidItems.includes(uniqueKey)) return;

             const relevantSchedules = schedules.filter(sch => 
                sch.subjectId === sub.id && 
                sch.classId === cls.id && 
                sch.status !== ScheduleStatus.OFF
            );
            
            const learned = relevantSchedules.reduce((acc, curr) => acc + curr.periodCount, 0);
            
            // Condition: Learned >= Total AND started (learned > 0)
            if (learned >= sub.totalPeriods && learned > 0) {
                 const teacherIds = Array.from(new Set(relevantSchedules.map(s => s.teacherId)));
                 const teacherNames = teacherIds.map(tid => teachers.find(t => t.id === tid)?.name || 'GV đã xóa').join(', ');

                 // Calculate End Date: Sort by date and take the last one
                 relevantSchedules.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
                 const lastSchedule = relevantSchedules[relevantSchedules.length - 1];
                 const endDate = lastSchedule ? format(parseLocal(lastSchedule.date), 'dd/MM/yyyy') : 'N/A';

                 results.push({
                     uniqueKey: uniqueKey,
                     subjectName: sub.name,
                     className: cls.name,
                     teacherName: teacherNames || "Chưa xác định",
                     totalPeriods: sub.totalPeriods,
                     endDate: endDate
                 });
            }
        });
    });
    
    return results;
  }, [subjects, schedules, classes, teachers, paidItems]);

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold text-gray-800">Tổng quan</h1>
      
      {/* Stats Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white p-4 rounded-xl shadow border border-gray-100 flex items-center space-x-4">
          <div className="p-3 bg-blue-100 text-blue-600 rounded-full">
            <BookOpen size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500">Tổng môn học</p>
            <p className="text-2xl font-bold">{subjects.length}</p>
          </div>
        </div>
        <div className="bg-white p-4 rounded-xl shadow border border-gray-100 flex items-center space-x-4">
          <div className="p-3 bg-green-100 text-green-600 rounded-full">
            <Users size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500">Tổng giáo viên</p>
            <p className="text-2xl font-bold">{teachers.length}</p>
          </div>
        </div>
        <div className="bg-white p-4 rounded-xl shadow border border-gray-100 flex items-center space-x-4">
          <div className="p-3 bg-yellow-100 text-yellow-600 rounded-full">
            <Calendar size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500">Lịch thi sắp tới</p>
            <p className="text-2xl font-bold">{upcomingExams.length}</p>
          </div>
        </div>
        <div className="bg-white p-4 rounded-xl shadow border border-gray-100 flex items-center space-x-4">
          <div className="p-3 bg-purple-100 text-purple-600 rounded-full">
            <Clock size={24} />
          </div>
          <div>
            <p className="text-sm text-gray-500">Hôm nay</p>
            <p className="text-lg font-bold">{format(currentTime, 'HH:mm dd/MM', { locale: vi })}</p>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Today's Schedule */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
          <h2 className="text-lg font-semibold mb-4 flex items-center">
            <Calendar className="mr-2 text-blue-500" size={20} />
            Lịch giảng dạy hôm nay
          </h2>
          {todaySchedules.length === 0 ? (
            <p className="text-gray-500 italic">Không có lịch học hôm nay.</p>
          ) : (
            <div className="space-y-3">
              {todaySchedules.map(s => {
                const subj = subjects.find(sub => sub.id === s.subjectId);
                const tea = teachers.find(t => t.id === s.teacherId);
                const cls = classes.find(c => c.id === s.classId);
                const computedStatus = determineStatus(s.date, s.startPeriod, s.status);
                
                return (
                  <div key={s.id} className="p-3 rounded-lg border-l-4 border-blue-500 bg-blue-50">
                    <div className="flex justify-between items-start">
                      <div>
                        <h3 className="font-bold text-gray-800">{subj?.name}</h3>
                        <p className="text-sm text-gray-600">GV: {tea?.name || 'GV đã xóa'}</p>
                        <p className="text-sm text-gray-500">Lớp: {cls?.name} - Phòng: {s.roomId}</p>
                      </div>
                      <div className="text-right">
                        <span className="inline-block px-2 py-1 text-xs font-semibold rounded bg-white border">
                          Tiết {s.startPeriod} - {s.startPeriod + s.periodCount - 1}
                        </span>
                        <div className="mt-1 text-xs text-gray-500 font-medium">
                           {computedStatus}
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Upcoming Exams */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
          <h2 className="text-lg font-semibold mb-4 flex items-center">
            <AlertTriangle className="mr-2 text-yellow-500" size={20} />
            Lịch thi sắp tới
          </h2>
          {upcomingExams.length === 0 ? (
            <p className="text-gray-500 italic">Chưa có lịch thi.</p>
          ) : (
            <div className="space-y-3">
              {upcomingExams.map(s => {
                const subj = subjects.find(sub => sub.id === s.subjectId);
                const cls = classes.find(c => c.id === s.classId);
                const isToday = isSameDay(parseLocal(s.date), today);
                
                return (
                  <div key={s.id} className={`flex items-center justify-between p-3 border rounded-lg shadow-sm hover:shadow-md transition ${isToday ? 'bg-yellow-50 border-yellow-200' : 'bg-white border-gray-100'}`}>
                    <div>
                      <p className="font-bold text-gray-800">{subj?.name} {isToday && <span className="text-red-500 text-xs">(Hôm nay)</span>}</p>
                      <p className="text-sm text-gray-500">{cls?.name}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-sm font-semibold text-blue-600">{format(parseLocal(s.date), 'dd/MM/yyyy')}</p>
                      <p className="text-xs text-gray-400">Phòng {s.roomId}</p>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {/* NEW: Completed Subjects */}
      <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
        <h2 className="text-lg font-semibold mb-4 flex items-center">
          <CheckCircle className="mr-2 text-green-500" size={20} />
          Môn học đã kết thúc
        </h2>
        
        {completedSubjects.length > 0 ? (
          <div className="overflow-x-auto">
            <table className="w-full text-left bg-white">
              <thead className="bg-gray-50 border-b">
                <tr>
                   <th className="p-3 text-sm font-semibold text-gray-600">Môn học</th>
                   <th className="p-3 text-sm font-semibold text-gray-600">Lớp</th>
                   <th className="p-3 text-sm font-semibold text-gray-600">Giáo viên</th>
                   <th className="p-3 text-sm font-semibold text-gray-600">Ngày kết thúc</th>
                   <th className="p-3 text-sm font-semibold text-gray-600 text-right">Tổng tiết</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {completedSubjects.map((item: any) => (
                  <tr key={item.uniqueKey} className="hover:bg-gray-50 transition-colors">
                    <td className="p-3 font-medium text-gray-800">{item.subjectName}</td>
                    <td className="p-3 text-gray-600 text-sm">{item.className}</td>
                    <td className="p-3 text-gray-600 text-sm">{item.teacherName}</td>
                    <td className="p-3 text-gray-600 text-sm">{item.endDate}</td>
                    <td className="p-3 text-gray-800 font-medium text-right">{item.totalPeriods}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="text-gray-500 italic py-6 text-center bg-gray-50 rounded border border-dashed border-gray-200">
             Chưa có môn học nào hoàn thành.
          </div>
        )}
      </div>
    </div>
  );
};

export default Dashboard;
