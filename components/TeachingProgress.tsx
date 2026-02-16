import React, { useState, useMemo, useEffect } from 'react';
import { useApp } from '../store/AppContext';
import { ScheduleStatus } from '../types';
import { Filter, BookOpen, User, Edit2, Save, X, Calendar, Clock, Users } from 'lucide-react';
import { parseLocal } from '../utils';
import { isSameDay, startOfDay, format } from 'date-fns';

// Interface for local metadata storage
interface SubjectMetadata {
    statusOverride?: 'upcoming' | 'completed' | 'in-progress' | 'auto';
    customStartDate?: string; // YYYY-MM-DD
    customEndDate?: string;   // YYYY-MM-DD
    customExamDate?: string;  // YYYY-MM-DD
}

const TeachingProgress: React.FC = () => {
  const { classes, subjects, schedules, teachers, updateSubject } = useApp();
  const [selectedClassId, setSelectedClassId] = useState<string>(classes[0]?.id || '');

  // NEW: State for Subject Metadata (Persisted in localStorage) replaces simple manualCompleted
  const [progressMetadata, setProgressMetadata] = useState<Record<string, SubjectMetadata>>(() => {
    try {
      const saved = localStorage.getItem('subject_progress_metadata');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Modal State
  const [editingItem, setEditingItem] = useState<{
      subjectId: string;
      subjectName: string;
      status: string;
      startDate: string;
      endDate: string;
      examDate: string;
      selectedTeacherId: string; // NEW
  } | null>(null);

  const saveMetadata = (key: string, data: SubjectMetadata) => {
      const newMeta = { ...progressMetadata, [key]: data };
      setProgressMetadata(newMeta);
      localStorage.setItem('subject_progress_metadata', JSON.stringify(newMeta));
  };

  const handleEditClick = (subject: any) => {
      // Convert display dates back to input format (YYYY-MM-DD) if possible, or keep empty
      const toInputDate = (dateStr: string) => {
          if (!dateStr || dateStr === '--/--/----' || dateStr === 'Chưa xếp') return '';
          const parts = dateStr.split('/');
          if (parts.length === 3) return `${parts[2]}-${parts[1]}-${parts[0]}`;
          return '';
      };

      // Find the teacher ID based on name to pre-select
      const matchedTeacher = teachers.find(t => t.name === subject.teacher1);

      setEditingItem({
          subjectId: subject.id,
          subjectName: subject.name,
          status: subject.status,
          startDate: toInputDate(subject.startDate),
          endDate: toInputDate(subject.endDate),
          examDate: toInputDate(subject.examDate),
          selectedTeacherId: matchedTeacher ? matchedTeacher.id : ''
      });
  };

  const handleSaveModal = () => {
      if (!editingItem || !selectedClassId) return;

      // VALIDATION: Check End Date > Start Date if Completed
      if (editingItem.status === 'completed' && editingItem.startDate && editingItem.endDate) {
          if (new Date(editingItem.startDate) > new Date(editingItem.endDate)) {
              alert("Ngày kết thúc phải lớn hơn hoặc bằng ngày bắt đầu.");
              return;
          }
      }

      const key = `${editingItem.subjectId}-${selectedClassId}`;
      
      let statusOverride: SubjectMetadata['statusOverride'] = 'auto';
      if (editingItem.status === 'completed') statusOverride = 'completed';
      else if (editingItem.status === 'upcoming') statusOverride = 'upcoming';
      else if (editingItem.status === 'in-progress') statusOverride = 'in-progress';

      const metadata: SubjectMetadata = {
          statusOverride: statusOverride,
          customStartDate: editingItem.startDate || undefined,
          customEndDate: editingItem.endDate || undefined,
          customExamDate: editingItem.examDate || undefined
      };

      // 1. Save Metadata
      saveMetadata(key, metadata);

      // 2. Save Teacher Assignment (Update Subject)
      const selectedTeacher = teachers.find(t => t.id === editingItem.selectedTeacherId);
      updateSubject(editingItem.subjectId, {
          teacher1: selectedTeacher ? selectedTeacher.name : '',
          phone1: selectedTeacher ? selectedTeacher.phone : ''
      });

      setEditingItem(null);
  };

  const currentClass = classes.find(c => c.id === selectedClassId);
  const today = startOfDay(new Date());

  // Calculate progress for all subjects in the selected class
  const progressData = useMemo(() => {
    if (!currentClass) return [];

    // 1. Get subjects for this class's major
    const classSubjects = subjects.filter(s => s.majorId === currentClass.majorId);

    // 2. Calculate stats per subject
    return classSubjects.map(sub => {
      const uniqueKey = `${sub.id}-${selectedClassId}`;
      const metadata = progressMetadata[uniqueKey] || {};

      const relevantSchedules = schedules.filter(sch => 
        sch.subjectId === sub.id && 
        sch.classId === selectedClassId && 
        sch.status !== ScheduleStatus.OFF
      );

      // Separate Class and Exam schedules
      const classSchedules = relevantSchedules
        .filter(s => s.type === 'class')
        .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
      
      const examSchedule = relevantSchedules.find(s => s.type === 'exam');

      // --- GROUP LOGIC ---
      // Identify distinct groups (e.g. "Nhóm 1", "Nhóm 2")
      const uniqueGroups = Array.from(new Set(classSchedules.map(s => s.group).filter((g): g is string => !!g && g !== ''))) as string[];
      
      let groupBreakdown: { name: string, learned: number, percentage: number }[] = [];
      let totalLearnedForStandard = 0; // Fallback if no groups
      let isAutoCompleted = false;

      if (uniqueGroups.length > 0) {
          // Logic: Subject has groups. 
          // Progress of Group X = (Common/Theory Periods) + (Group X Practice Periods)
          
          // 1. Calculate Common (Theory) periods - those without a group
          const commonSchedules = classSchedules.filter(s => !s.group);
          const commonRealized = commonSchedules.reduce((acc, curr) => parseLocal(curr.date) <= today ? acc + curr.periodCount : acc, 0);

          // 2. Calculate per group
          groupBreakdown = uniqueGroups.sort().map(gName => {
               const gSchedules = classSchedules.filter(s => s.group === gName);
               const gRealized = gSchedules.reduce((acc, curr) => parseLocal(curr.date) <= today ? acc + curr.periodCount : acc, 0);
               const totalForGroup = commonRealized + gRealized;
               
               return {
                   name: gName,
                   learned: totalForGroup,
                   percentage: Math.min(100, Math.round((totalForGroup / sub.totalPeriods) * 100))
               };
          });

          // Auto Complete only if ALL groups are finished
          isAutoCompleted = groupBreakdown.every(g => g.percentage >= 100);

          // For sorting/display, use the MIN percentage as the "Subject Percentage" to be conservative
          totalLearnedForStandard = Math.min(...groupBreakdown.map(g => g.learned)); 
      } else {
          // Standard Logic (No groups)
          totalLearnedForStandard = classSchedules.reduce((acc, curr) => {
              const sDate = parseLocal(curr.date);
              if (sDate <= today) return acc + curr.periodCount;
              return acc;
          }, 0);
          isAutoCompleted = totalLearnedForStandard >= sub.totalPeriods;
      }

      const totalPeriods = sub.totalPeriods;
      const percentage = Math.min(100, Math.round((totalLearnedForStandard / totalPeriods) * 100));


      // --- DATE CALCULATION (Prioritize Metadata) ---
      let startDate = '--/--/----';
      let endDate = '--/--/----';
      let examDate = 'Chưa xếp';

      // Start Date
      if (metadata.customStartDate) {
          startDate = format(parseLocal(metadata.customStartDate), 'dd/MM/yyyy');
      } else if (classSchedules.length > 0) {
          startDate = format(parseLocal(classSchedules[0].date), 'dd/MM/yyyy');
      }

      // End Date
      if (metadata.customEndDate) {
          endDate = format(parseLocal(metadata.customEndDate), 'dd/MM/yyyy');
      } else if (classSchedules.length > 0) {
          endDate = format(parseLocal(classSchedules[classSchedules.length - 1].date), 'dd/MM/yyyy');
      }

      // Exam Date
      if (metadata.customExamDate) {
           examDate = format(parseLocal(metadata.customExamDate), 'dd/MM/yyyy');
      } else if (examSchedule) {
           examDate = format(parseLocal(examSchedule.date), 'dd/MM/yyyy');
      }

      // --- STATUS CALCULATION ---
      const hasScheduleToday = relevantSchedules.some(s => isSameDay(parseLocal(s.date), today));
      const hasSchedulePast = relevantSchedules.some(s => parseLocal(s.date) < today);

      let status: 'upcoming' | 'in-progress' | 'completed' = 'upcoming';

      // 2. Apply Override if exists
      if (metadata.statusOverride && metadata.statusOverride !== 'auto') {
          status = metadata.statusOverride;
      } else {
          // Auto Logic
          if (isAutoCompleted) status = 'completed';
          else if (hasScheduleToday || hasSchedulePast) status = 'in-progress';
          else status = 'upcoming';
      }

      return {
        ...sub,
        learnedPeriods: totalLearnedForStandard, // Represents conservative progress
        groupBreakdown, // NEW: Detailed group stats
        percentage, 
        status,
        startDate,
        endDate,
        examDate,
        isCustomized: !!metadata.statusOverride || !!metadata.customStartDate
      };
    }).sort((a, b) => {
        // Sort order: In Progress -> Upcoming -> Completed
        const order = { 'in-progress': 1, 'upcoming': 2, 'completed': 3 };
        return order[a.status] - order[b.status];
    });
  }, [subjects, schedules, classes, selectedClassId, currentClass, today, progressMetadata]);

  // Summary Counts
  const summary = {
    total: progressData.length,
    completed: progressData.filter(p => p.status === 'completed').length,
    inProgress: progressData.filter(p => p.status === 'in-progress').length,
    upcoming: progressData.filter(p => p.status === 'upcoming').length,
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'completed': return 'bg-green-100 text-green-700 border-green-200';
      case 'in-progress': return 'bg-blue-100 text-blue-700 border-blue-200';
      default: return 'bg-gray-100 text-gray-600 border-gray-200';
    }
  };

  const getStatusLabel = (status: string) => {
    switch (status) {
      case 'completed': return 'Đã hoàn thành';
      case 'in-progress': return 'Đang học';
      default: return 'Sắp học';
    }
  };

  const getProgressBarColor = (status: string) => {
     switch (status) {
      case 'completed': return 'bg-green-500';
      case 'in-progress': return 'bg-blue-500';
      default: return 'bg-gray-300';
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <h1 className="text-2xl font-bold text-gray-800">Tiến độ giảng dạy</h1>
        
        <div className="flex items-center gap-3 bg-white p-2 rounded shadow-sm border">
            <Filter size={20} className="text-blue-600" />
            <span className="font-semibold text-gray-700 whitespace-nowrap text-sm">Chọn lớp xem:</span>
            <select 
                className="border-none outline-none bg-transparent font-medium text-gray-800 min-w-[200px]"
                value={selectedClassId}
                onChange={(e) => setSelectedClassId(e.target.value)}
            >
                {classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
        </div>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-200">
           <div className="text-gray-500 text-xs font-semibold uppercase">Tổng môn học</div>
           <div className="text-2xl font-bold text-gray-800 mt-1">{summary.total}</div>
        </div>
        <div className="bg-blue-50 p-4 rounded-xl shadow-sm border border-blue-100">
           <div className="text-blue-600 text-xs font-semibold uppercase">Đang học</div>
           <div className="text-2xl font-bold text-blue-700 mt-1">{summary.inProgress}</div>
        </div>
        <div className="bg-green-50 p-4 rounded-xl shadow-sm border border-green-100">
           <div className="text-green-600 text-xs font-semibold uppercase">Đã hoàn thành</div>
           <div className="text-2xl font-bold text-green-700 mt-1">{summary.completed}</div>
        </div>
        <div className="bg-gray-50 p-4 rounded-xl shadow-sm border border-gray-200">
           <div className="text-gray-500 text-xs font-semibold uppercase">Sắp học</div>
           <div className="text-2xl font-bold text-gray-600 mt-1">{summary.upcoming}</div>
        </div>
      </div>

      {/* Progress List View */}
      <div className="bg-white rounded-xl shadow border border-gray-200 overflow-hidden">
          <div className="p-4 border-b bg-gray-50 flex justify-between items-center">
             <h3 className="font-bold text-gray-700">Chi tiết môn học</h3>
             <span className="text-sm text-gray-500">Tổng số: {progressData.length} môn</span>
          </div>

          {/* Table Header (Visible on Desktop) */}
          <div className="hidden md:grid grid-cols-12 gap-4 px-4 py-3 bg-gray-100 border-b text-xs font-bold text-gray-500 uppercase tracking-wider">
             <div className="col-span-4">Môn học / Giáo viên</div>
             <div className="col-span-3">Lịch trình & Thi</div>
             <div className="col-span-3">Tiến độ</div>
             <div className="col-span-2 text-center">Trạng thái</div>
          </div>
          
          <div className="divide-y divide-gray-100">
             {progressData.map(subject => (
                 <div key={subject.id} className="p-4 hover:bg-gray-50 transition grid grid-cols-1 md:grid-cols-12 gap-4 items-center">
                     {/* Column 1: Icon & Name & Teacher (4 cols) */}
                     <div className="md:col-span-4 flex items-start gap-3">
                        <div className={`p-2 rounded-lg flex-shrink-0 ${subject.status === 'in-progress' ? 'bg-blue-100 text-blue-600' : 'bg-gray-100 text-gray-400'}`}>
                             <BookOpen size={20} />
                        </div>
                        <div>
                             <h4 className="font-bold text-gray-800 flex items-center">
                                 {subject.name}
                                 {subject.isCustomized && <span className="ml-2 w-2 h-2 rounded-full bg-orange-400" title="Đã chỉnh sửa thủ công"></span>}
                             </h4>
                             <div className="flex items-center text-sm text-gray-500 mt-1">
                                <User size={14} className="mr-1" />
                                {subject.teacher1 || <span className="italic text-gray-400">Chưa phân công</span>}
                             </div>
                        </div>
                     </div>

                     {/* Column 2: Dates (3 cols) */}
                     <div className="md:col-span-3 text-xs space-y-1.5 text-gray-600 pl-11 md:pl-0">
                         <div className="flex items-center gap-2">
                             <span className="w-16 font-semibold text-gray-500">Bắt đầu:</span>
                             <span className="font-medium">{subject.startDate}</span>
                         </div>
                         <div className="flex items-center gap-2">
                             <span className="w-16 font-semibold text-gray-500">Kết thúc:</span>
                             <span className="font-medium">{subject.endDate}</span>
                         </div>
                         <div className="flex items-center gap-2">
                             <span className="w-16 font-semibold text-gray-500">Thi:</span>
                             <span className={`font-medium ${subject.examDate !== 'Chưa xếp' ? 'text-blue-600' : 'text-gray-400 italic'}`}>
                                 {subject.examDate}
                             </span>
                         </div>
                     </div>

                     {/* Column 3: Progress Bar (3 cols) */}
                     <div className="md:col-span-3 w-full px-2">
                        {/* Check if we have group breakdown */}
                        {subject.groupBreakdown && subject.groupBreakdown.length > 0 ? (
                            <div className="space-y-2">
                                <div className="text-[10px] font-bold text-blue-600 uppercase flex items-center mb-1">
                                    <Users size={10} className="mr-1" /> Tiến độ theo nhóm
                                </div>
                                {subject.groupBreakdown.map(gb => (
                                    <div key={gb.name}>
                                        <div className="flex justify-between text-[10px] mb-0.5">
                                            <span className="font-semibold text-gray-600">{gb.name}</span>
                                            <span className="text-gray-500">{gb.learned}/{subject.totalPeriods} ({gb.percentage}%)</span>
                                        </div>
                                        <div className="w-full bg-gray-200 rounded-full h-1.5">
                                            <div className={`h-1.5 rounded-full ${getProgressBarColor(subject.status)}`} style={{width: `${gb.percentage}%`}}></div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                        ) : (
                            // Standard Single Bar
                            <>
                                <div className="flex justify-between text-xs mb-1.5 font-medium">
                                    <span className="text-gray-600">{subject.learnedPeriods} / {subject.totalPeriods} tiết</span>
                                    <span className={`${subject.status === 'completed' ? 'text-green-600' : subject.status === 'in-progress' ? 'text-blue-600' : 'text-gray-400'}`}>
                                        {subject.percentage}%
                                    </span>
                                </div>
                                <div className="w-full bg-gray-200 rounded-full h-2.5">
                                    <div 
                                        className={`h-2.5 rounded-full transition-all duration-500 ${getProgressBarColor(subject.status)}`} 
                                        style={{ width: `${subject.percentage}%` }}
                                    ></div>
                                </div>
                            </>
                        )}
                     </div>

                     {/* Column 4: Status Badge & Edit Button (2 cols) */}
                     <div className="md:col-span-2 w-full flex flex-row md:flex-col lg:flex-row justify-between md:justify-center items-center gap-3">
                         <span className={`px-2 py-1 rounded text-xs font-bold border whitespace-nowrap flex items-center gap-1 ${getStatusColor(subject.status)}`}>
                            {getStatusLabel(subject.status)}
                         </span>
                         
                         <button
                            type="button"
                            onClick={() => handleEditClick(subject)}
                            className="px-2 py-1 bg-white text-gray-600 border border-gray-300 rounded hover:bg-gray-100 hover:text-blue-600 transition-colors shadow-sm text-xs font-medium flex items-center"
                            title="Điều chỉnh lịch trình và trạng thái"
                         >
                            <Edit2 size={12} className="mr-1" /> Chỉnh sửa
                         </button>
                     </div>
                 </div>
             ))}

             {progressData.length === 0 && (
                <div className="p-8 text-center text-gray-400 italic">
                    Chưa có môn học nào thuộc ngành của lớp này.
                </div>
             )}
          </div>
      </div>

      {/* EDIT MODAL */}
      {editingItem && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
             <div className="bg-white rounded-xl shadow-xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
                 <div className="p-4 border-b bg-gray-50 flex justify-between items-center">
                    <div>
                        <h3 className="font-bold text-gray-800">Cập nhật môn học</h3>
                        <p className="text-xs text-blue-600 font-medium">{editingItem.subjectName}</p>
                    </div>
                    <button onClick={() => setEditingItem(null)} className="text-gray-400 hover:text-red-500">
                        <X size={20} />
                    </button>
                 </div>
                 
                 <div className="p-6 space-y-4">
                     <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Trạng thái môn học</label>
                        <select 
                            className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                            value={editingItem.status}
                            onChange={(e) => {
                                const newStatus = e.target.value;
                                let updates: any = { status: newStatus };
                                // If user chooses 'upcoming', clear existing dates
                                if (newStatus === 'upcoming') {
                                    updates.startDate = '';
                                    updates.endDate = '';
                                    updates.examDate = '';
                                }
                                setEditingItem({...editingItem, ...updates});
                            }}
                        >
                            <option value="auto">Tự động (Theo tiến độ thực tế)</option>
                            <option value="upcoming">Sắp học (Chưa triển khai)</option>
                            <option value="in-progress">Đang học</option>
                            <option value="completed">Đã hoàn thành</option>
                        </select>
                        <p className="text-[10px] text-gray-500 mt-1">
                            * "Tự động": Hệ thống tự tính toán dựa trên số tiết đã dạy và ngày tháng.
                        </p>
                     </div>

                     <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Phân công giáo viên</label>
                        <select 
                            className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                            value={editingItem.selectedTeacherId}
                            onChange={(e) => setEditingItem({...editingItem, selectedTeacherId: e.target.value})}
                        >
                            <option value="">-- Chưa phân công --</option>
                            {teachers.map(t => (
                                <option key={t.id} value={t.id}>{t.name}</option>
                            ))}
                        </select>
                     </div>

                     <div className="grid grid-cols-2 gap-4">
                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">Ngày bắt đầu</label>
                            <input 
                                type="date"
                                className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none"
                                value={editingItem.startDate}
                                onChange={(e) => setEditingItem({...editingItem, startDate: e.target.value})}
                            />
                        </div>
                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">Ngày kết thúc</label>
                            <input 
                                type="date"
                                className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none"
                                value={editingItem.endDate}
                                onChange={(e) => setEditingItem({...editingItem, endDate: e.target.value})}
                            />
                        </div>
                     </div>

                     <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Ngày thi (Dự kiến/Chính thức)</label>
                        <input 
                            type="date"
                            className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none"
                            value={editingItem.examDate}
                            onChange={(e) => setEditingItem({...editingItem, examDate: e.target.value})}
                        />
                     </div>
                 </div>

                 <div className="p-4 border-t bg-gray-50 flex justify-end gap-2">
                    <button onClick={() => setEditingItem(null)} className="px-4 py-2 text-gray-600 hover:bg-gray-200 rounded text-sm">Hủy</button>
                    <button onClick={handleSaveModal} className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 flex items-center gap-2 text-sm shadow">
                        <Save size={16} /> Lưu thay đổi
                    </button>
                 </div>
             </div>
        </div>
      )}
    </div>
  );
};

export default TeachingProgress;