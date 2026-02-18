
import React, { createContext, useContext, useState, useEffect } from 'react';
import { Teacher, Subject, ClassEntity, ScheduleItem, Major, ScheduleStatus, Student, AppState, DocumentItem, ExportTemplate, Holiday } from '../types';
import { generateId } from '../utils';

interface AppContextType extends AppState {
  addTeacher: (t: Omit<Teacher, 'id'>) => void;
  updateTeacher: (id: string, t: Partial<Teacher>) => void;
  deleteTeacher: (id: string) => void;
  importTeachers: (teachers: Omit<Teacher, 'id'>[]) => void;
  addSubject: (s: Omit<Subject, 'id'>) => void;
  updateSubject: (id: string, s: Partial<Subject>) => void;
  deleteSubject: (id: string) => void;
  importSubjects: (subjects: Omit<Subject, 'id'>[]) => void;
  addSchedule: (s: Omit<ScheduleItem, 'id' | 'status'>) => void;
  updateSchedule: (id: string, s: Partial<ScheduleItem>) => void;
  deleteSchedule: (id: string) => void;
  addClass: (c: Omit<ClassEntity, 'id'>) => void;
  updateClass: (id: string, c: Partial<ClassEntity>) => void;
  deleteClass: (id: string) => void;
  importClasses: (classes: Omit<ClassEntity, 'id'>[]) => void;
  addStudent: (s: Omit<Student, 'id'>) => void;
  updateStudent: (id: string, s: Partial<Student>) => void;
  deleteStudent: (id: string) => void;
  importStudents: (students: Omit<Student, 'id'>[]) => void;
  addDocument: (d: Omit<DocumentItem, 'id'>) => void;
  deleteDocument: (id: string) => void;
  addTemplate: (t: Omit<ExportTemplate, 'id'>) => void;
  deleteTemplate: (id: string) => void;
  addHoliday: (h: Omit<Holiday, 'id'>) => void;
  updateHoliday: (id: string, h: Partial<Holiday>) => void;
  deleteHoliday: (id: string) => void;
  loadData: (data: AppState) => void;
  resetData: () => void;
}

const AppContext = createContext<AppContextType | undefined>(undefined);

const INITIAL_DATA: AppState = {
  teachers: [
    { id: '1', title: 'Thầy', name: 'Nguyễn Văn Thái', phone: '0901234567', bank: 'VCB', accountNumber: '123456', mainSubject: '1', ratePerPeriod: 100000 },
    { id: '2', title: 'Cô', name: 'Trần Thị Hà', phone: '0909876543', bank: 'ACB', accountNumber: '654321', mainSubject: '2', ratePerPeriod: 120000 },
    { id: '3', title: 'Thầy', name: 'Lê Văn Kha', phone: '0912345678', bank: 'Tech', accountNumber: '888888', mainSubject: '3', ratePerPeriod: 110000 },
  ],
  subjects: [
    { id: '1', name: 'Đo lường và TBĐ', majorId: '2', totalPeriods: 30 }, 
    { id: '2', name: 'Lập trình C++', majorId: '4', totalPeriods: 45 },
    { id: '3', name: 'Nguyên lý kế toán', majorId: '1', totalPeriods: 60 },
    { id: '4', name: 'Khí cụ điện', majorId: '2', totalPeriods: 45 },
    { id: '5', name: 'Mạch điện tử', majorId: '3', totalPeriods: 60 },
    // Môn chung
    { id: '6', name: 'Giáo dục chính trị', majorId: 'common', totalPeriods: 30, isShared: true },
    { id: '7', name: 'Tiếng Anh cơ bản', majorId: 'common', totalPeriods: 45, isShared: true },
    { id: '8', name: 'Giáo dục thể chất', majorId: 'common', totalPeriods: 30, isShared: true },
    // Môn văn hóa
    { id: '9', name: 'Toán 10', majorId: 'culture', totalPeriods: 45, isShared: true },
    { id: '10', name: 'Ngữ văn 10', majorId: 'culture', totalPeriods: 45, isShared: true },
  ],
  classes: [
    { id: '1', name: 'Điện Công Nghiệp (25DC2H8)', studentCount: 40, majorId: '2', schoolYear: '2023-2026' },
    { id: '2', name: 'Kế toán K15', studentCount: 35, majorId: '1', schoolYear: '2023-2026' },
  ],
  students: [
    { id: '1', studentCode: 'SV001', classId: '1', name: 'Nguyễn Văn A', dob: '2005-01-15', pob: 'Hà Nội', fatherName: 'Nguyễn Văn B', motherName: 'Lê Thị C', phone: '0987654321', status: 'studying' },
    { id: '2', studentCode: 'SV002', classId: '1', name: 'Trần Thị B', dob: '2005-05-20', pob: 'Nam Định', fatherName: 'Trần Văn D', motherName: 'Phạm Thị E', phone: '0912345678', status: 'studying' },
  ],
  majors: [
    { id: 'common', name: 'Môn chung' },
    { id: 'culture', name: 'Văn hóa' },
    { id: 'culture_8', name: 'Văn hóa 8 môn' }, // Added new major category
    { id: '1', name: 'Kế toán Doanh nghiệp' },
    { id: '2', name: 'Điện công nghiệp' },
    { id: '3', name: 'Điện - điện tử' },
    { id: '4', name: 'Công nghệ thông tin' },
  ],
  schedules: [],
  documents: [],
  templates: [],
  holidays: []
};

export const AppProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  // Load from local storage or use initial
  const [state, setState] = useState<AppState>(() => {
    try {
      const saved = localStorage.getItem('eduScheduleData');
      if (saved) {
          const parsed = JSON.parse(saved);
          // Ensure arrays exist for older saved data (Compatibility check)
          return {
              ...INITIAL_DATA,
              ...parsed,
              teachers: parsed.teachers || [],
              subjects: parsed.subjects || [],
              classes: parsed.classes || [],
              students: parsed.students || [],
              schedules: parsed.schedules || [],
              majors: parsed.majors || INITIAL_DATA.majors,
              documents: parsed.documents || [],
              templates: parsed.templates || [],
              holidays: parsed.holidays || []
          };
      }
      return INITIAL_DATA;
    } catch (e) {
      console.error("Failed to load data from localStorage", e);
      return INITIAL_DATA;
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem('eduScheduleData', JSON.stringify(state));
    } catch (error) {
      console.error("Storage Limit Exceeded", error);
      alert("Cảnh báo: Bộ nhớ trình duyệt đã đầy! Dữ liệu mới chưa được lưu. Vui lòng xóa bớt tài liệu hoặc sao lưu và reset hệ thống.");
    }
  }, [state]);

  const addTeacher = (t: Omit<Teacher, 'id'>) => {
    setState(prev => ({ ...prev, teachers: [...prev.teachers, { ...t, id: generateId() }] }));
  };
  const updateTeacher = (id: string, t: Partial<Teacher>) => {
    setState(prev => ({
        ...prev,
        teachers: prev.teachers.map(tea => tea.id === id ? { ...tea, ...t } : tea)
    }))
  }
  const deleteTeacher = (id: string) => {
    setState(prev => ({ ...prev, teachers: prev.teachers.filter(t => t.id !== id) }));
  };
  const importTeachers = (newTeachers: Omit<Teacher, 'id'>[]) => {
      setState(prev => ({
          ...prev,
          teachers: [...prev.teachers, ...newTeachers.map(t => ({...t, id: generateId()}))]
      }));
  }

  const addSubject = (s: Omit<Subject, 'id'>) => {
    setState(prev => ({ ...prev, subjects: [...prev.subjects, { ...s, id: generateId() }] }));
  };
  const updateSubject = (id: string, s: Partial<Subject>) => {
    setState(prev => ({
        ...prev,
        subjects: prev.subjects.map(sub => sub.id === id ? { ...sub, ...s } : sub)
    }))
  };
  const deleteSubject = (id: string) => {
    setState(prev => ({ ...prev, subjects: prev.subjects.filter(s => s.id !== id) }));
  };
  const importSubjects = (newSubjects: Omit<Subject, 'id'>[]) => {
      setState(prev => ({
          ...prev,
          subjects: [...prev.subjects, ...newSubjects.map(s => ({...s, id: generateId()}))]
      }));
  }

  const addClass = (c: Omit<ClassEntity, 'id'>) => {
    setState(prev => ({ ...prev, classes: [...prev.classes, { ...c, id: generateId() }] }));
  };
  const updateClass = (id: string, c: Partial<ClassEntity>) => {
    setState(prev => ({
        ...prev,
        classes: prev.classes.map(cls => cls.id === id ? { ...cls, ...c } : cls)
    }))
  };
  const deleteClass = (id: string) => {
    setState(prev => ({ ...prev, classes: prev.classes.filter(c => c.id !== id) }));
  };
  const importClasses = (newClasses: Omit<ClassEntity, 'id'>[]) => {
      setState(prev => ({
          ...prev,
          classes: [...prev.classes, ...newClasses.map(c => ({...c, id: generateId()}))]
      }));
  }

  const addStudent = (s: Omit<Student, 'id'>) => {
    setState(prev => ({ ...prev, students: [...prev.students, { ...s, id: generateId() }] }));
  };
  const updateStudent = (id: string, s: Partial<Student>) => {
     setState(prev => ({
        ...prev,
        students: prev.students.map(stu => stu.id === id ? { ...stu, ...s } : stu)
    }))
  };
  const deleteStudent = (id: string) => {
    setState(prev => ({ ...prev, students: prev.students.filter(s => s.id !== id) }));
  };
  const importStudents = (newStudents: Omit<Student, 'id'>[]) => {
      setState(prev => ({
          ...prev,
          students: [...prev.students, ...newStudents.map(s => ({...s, id: generateId()}))]
      }));
  }

  const addSchedule = (s: Omit<ScheduleItem, 'id' | 'status'>) => {
    const newItem: ScheduleItem = { ...s, id: generateId(), status: ScheduleStatus.PENDING };
    setState(prev => ({ ...prev, schedules: [...prev.schedules, newItem] }));
  };

  const updateSchedule = (id: string, s: Partial<ScheduleItem>) => {
    setState(prev => {
        const updatedList = prev.schedules.map(item => item.id === id ? { ...item, ...s } : item);
        return { ...prev, schedules: updatedList };
    });
  };

  const deleteSchedule = (id: string) => {
    setState(prev => ({ ...prev, schedules: prev.schedules.filter(s => s.id !== id) }));
  };

  // Document actions
  const addDocument = (d: Omit<DocumentItem, 'id'>) => {
    setState(prev => ({ ...prev, documents: [...prev.documents, { ...d, id: generateId() }] }));
  };

  const deleteDocument = (id: string) => {
    setState(prev => ({ ...prev, documents: prev.documents.filter(d => d.id !== id) }));
  };

  // Template actions
  const addTemplate = (t: Omit<ExportTemplate, 'id'>) => {
    setState(prev => ({ ...prev, templates: [...prev.templates, { ...t, id: generateId() }] }));
  };
  
  const deleteTemplate = (id: string) => {
     setState(prev => ({ ...prev, templates: prev.templates.filter(t => t.id !== id) }));
  };

  // Holiday actions
  const addHoliday = (h: Omit<Holiday, 'id'>) => {
    setState(prev => ({ ...prev, holidays: [...prev.holidays, { ...h, id: generateId() }] }));
  };

  const updateHoliday = (id: string, h: Partial<Holiday>) => {
    setState(prev => ({
        ...prev,
        holidays: prev.holidays.map(item => item.id === id ? { ...item, ...h } : item)
    }));
  };

  const deleteHoliday = (id: string) => {
    setState(prev => ({ ...prev, holidays: prev.holidays.filter(h => h.id !== id) }));
  };

  // NEW: Load entire state (Restore) - Robust Version
  const loadData = (data: AppState) => {
      if (data && typeof data === 'object') {
          // Merge with INITIAL_DATA to ensure all fields (especially arrays like templates/documents) exist
          // regardless of the backup version.
          const robustData: AppState = {
              ...INITIAL_DATA,
              ...data,
              // Force Arrays if they are null/undefined in backup
              teachers: Array.isArray(data.teachers) ? data.teachers : [],
              subjects: Array.isArray(data.subjects) ? data.subjects : [],
              classes: Array.isArray(data.classes) ? data.classes : [],
              students: Array.isArray(data.students) ? data.students : [],
              schedules: Array.isArray(data.schedules) ? data.schedules : [],
              majors: Array.isArray(data.majors) ? data.majors : INITIAL_DATA.majors,
              documents: Array.isArray(data.documents) ? data.documents : [],
              templates: Array.isArray(data.templates) ? data.templates : [],
              holidays: Array.isArray(data.holidays) ? data.holidays : [],
          };
          
          setState(robustData);
          alert('Khôi phục dữ liệu thành công!');
      } else {
          alert('File dữ liệu không hợp lệ!');
      }
  };

  // NEW: Delete all data (Empty state, but keep Majors config)
  const resetData = () => {
    setState({
        teachers: [],
        subjects: [],
        classes: [],
        students: [],
        schedules: [],
        majors: INITIAL_DATA.majors, // Keep structural configuration
        documents: [],
        templates: [],
        holidays: []
    });
    // Clear auxiliary storage
    try {
        localStorage.removeItem('paid_completed_subjects');
        localStorage.removeItem('manual_completed_subjects');
        localStorage.removeItem('subject_progress_metadata');
    } catch (e) {
        console.error(e);
    }
  };

  return (
    <AppContext.Provider value={{ 
      ...state, 
      addTeacher, updateTeacher, deleteTeacher, importTeachers,
      addSubject, updateSubject, deleteSubject, importSubjects,
      addSchedule, updateSchedule, deleteSchedule, 
      addClass, updateClass, deleteClass, importClasses,
      addStudent, updateStudent, deleteStudent, importStudents,
      addDocument, deleteDocument,
      addTemplate, deleteTemplate,
      addHoliday, updateHoliday, deleteHoliday,
      loadData, resetData
    }}>
      {children}
    </AppContext.Provider>
  );
};

export const useApp = () => {
  const context = useContext(AppContext);
  if (!context) throw new Error('useApp must be used within AppProvider');
  return context;
};
