import React, { useState, useRef, useMemo } from 'react';
import { useApp } from '../store/AppContext';
import { Teacher, Subject, ClassEntity } from '../types';
import { Plus, Trash2, Edit2, Save, X, Filter, Search, Phone, Upload, HelpCircle, Users, ArrowUpDown, ArrowUp, ArrowDown } from 'lucide-react';
import XLSX from 'xlsx';

const Management: React.FC = () => {
  const { 
    teachers, subjects, majors, classes,
    addTeacher, updateTeacher, deleteTeacher, importTeachers,
    addSubject, updateSubject, deleteSubject, importSubjects,
    addClass, updateClass, deleteClass, importClasses
  } = useApp();
  
  const [activeTab, setActiveTab] = useState<'teachers' | 'subjects' | 'classes'>('teachers');

  const [newTeacher, setNewTeacher] = useState<Partial<Teacher>>({ title: 'Thầy' });
  const [editingTeacherId, setEditingTeacherId] = useState<string | null>(null);
  const [teacherSearch, setTeacherSearch] = useState('');
  const [teacherSort, setTeacherSort] = useState<'asc' | 'desc' | 'default'>('default');

  const [newSubject, setNewSubject] = useState<Partial<Subject>>({});
  const [editingSubjectId, setEditingSubjectId] = useState<string | null>(null);
  const [filterMajorId, setFilterMajorId] = useState<string>('');
  const [subjectSearch, setSubjectSearch] = useState('');

  const [newClass, setNewClass] = useState<Partial<ClassEntity>>({});
  const [editingClassId, setEditingClassId] = useState<string | null>(null);
  const [classSearch, setClassSearch] = useState('');
  const [classSortMajor, setClassSortMajor] = useState<'asc' | 'desc' | 'default'>('default'); // NEW State for class sorting

  const fileInputTeacherRef = useRef<HTMLInputElement>(null);
  const fileInputSubjectRef = useRef<HTMLInputElement>(null);
  const fileInputClassRef = useRef<HTMLInputElement>(null);

  // Teacher Sorting Logic
  const sortedTeachers = useMemo(() => {
    let list = teachers.filter(t => t.name.toLowerCase().includes(teacherSearch.toLowerCase()));

    if (teacherSort === 'default') return list;

    return [...list].sort((a, b) => {
        // Helper to split name into [LastName+MiddleName, FirstName]
        const getNameParts = (fullName: string) => {
            const parts = fullName.trim().split(' ');
            const firstName = parts.pop() || '';
            const lastName = parts.join(' ');
            return { firstName: firstName.toLowerCase(), lastName: lastName.toLowerCase() };
        };

        const nameA = getNameParts(a.name);
        const nameB = getNameParts(b.name);

        // 1. Compare First Name
        const firstCompare = nameA.firstName.localeCompare(nameB.firstName, 'vi');
        if (firstCompare !== 0) {
            return teacherSort === 'asc' ? firstCompare : -firstCompare;
        }

        // 2. Compare Last Name (if First Names are identical)
        const lastCompare = nameA.lastName.localeCompare(nameB.lastName, 'vi');
        return teacherSort === 'asc' ? lastCompare : -lastCompare;
    });
  }, [teachers, teacherSearch, teacherSort]);

  const toggleTeacherSort = () => {
      if (teacherSort === 'default') setTeacherSort('asc');
      else if (teacherSort === 'asc') setTeacherSort('desc');
      else setTeacherSort('default');
  };

  // Class Sorting Logic (NEW)
  const sortedClasses = useMemo(() => {
      let list = classes.filter(c => c.name.toLowerCase().includes(classSearch.toLowerCase()));

      if (classSortMajor === 'default') return list;

      return [...list].sort((a, b) => {
          const majorA = majors.find(m => m.id === a.majorId)?.name || '';
          const majorB = majors.find(m => m.id === b.majorId)?.name || '';
          
          const compare = majorA.localeCompare(majorB, 'vi');
          return classSortMajor === 'asc' ? compare : -compare;
      });
  }, [classes, classSearch, classSortMajor, majors]);

  const toggleClassSortMajor = () => {
      if (classSortMajor === 'default') setClassSortMajor('asc');
      else if (classSortMajor === 'asc') setClassSortMajor('desc');
      else setClassSortMajor('default');
  };

  // Handlers
  const handleSaveTeacher = () => {
    if (newTeacher.name) {
      const teacherData = {
          title: newTeacher.title || 'Thầy',
          name: newTeacher.name,
          phone: newTeacher.phone || '',
          bank: newTeacher.bank || '',
          accountNumber: newTeacher.accountNumber || '',
          mainSubject: newTeacher.mainSubject || '',
          ratePerPeriod: newTeacher.ratePerPeriod || 0
      };

      if (editingTeacherId) {
        updateTeacher(editingTeacherId, teacherData);
        setEditingTeacherId(null);
      } else {
        addTeacher(teacherData as Teacher);
      }
      setNewTeacher({ title: 'Thầy' });
    }
  };

  const handleEditTeacher = (t: Teacher) => {
    setNewTeacher(t);
    setEditingTeacherId(t.id);
  };

  const handleCancelTeacher = () => {
    setNewTeacher({ title: 'Thầy' });
    setEditingTeacherId(null);
  }

  const handleTeacherFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      // Reset immediately to prevent input freezing
      e.target.value = '';

      const reader = new FileReader();
      reader.onload = (evt) => {
          try {
              const bstr = evt.target?.result;
              const wb = XLSX.read(bstr, { type: 'binary' });
              const wsname = wb.SheetNames[0];
              const ws = wb.Sheets[wsname];
              const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
              
              if (!data || data.length < 2) {
                  setTimeout(() => alert("File Excel rỗng hoặc không đúng định dạng!"), 100);
                  return;
              }

              // Remove header
              const rows = data.slice(1) as any[];
              
              // Map columns: 0: Title, 1: Name, 2: Phone, 3: Bank, 4: Account, 5: Rate, 6: MainSubject
              const newTeachers: Omit<Teacher, 'id'>[] = rows.map(row => ({
                  title: row[0] || 'Thầy',
                  name: row[1] || '',
                  phone: row[2] ? String(row[2]) : '',
                  bank: row[3] ? String(row[3]) : '',
                  accountNumber: row[4] ? String(row[4]) : '',
                  ratePerPeriod: row[5] ? Number(row[5]) : 0,
                  mainSubject: row[6] ? String(row[6]) : '', // Mapped Main Subject
              })).filter(t => t.name);

              if (newTeachers.length > 0) {
                  importTeachers(newTeachers);
                  setTimeout(() => alert(`Đã nhập thành công ${newTeachers.length} giáo viên.`), 100);
              } else {
                  setTimeout(() => alert("Không tìm thấy dữ liệu hợp lệ."), 100);
              }
          } catch (error) {
              console.error(error);
              setTimeout(() => alert("Lỗi khi đọc file Excel."), 100);
          }
      };
      reader.readAsBinaryString(file);
  };

  const handleSaveSubject = () => {
    const selectedMajor = newSubject.majorId || filterMajorId;

    if (newSubject.name && newSubject.totalPeriods && selectedMajor) {
       const subjectData: any = {
           name: newSubject.name,
           majorId: selectedMajor,
           totalPeriods: newSubject.totalPeriods || 0,
           isShared: newSubject.isShared || false,
           teacher1: newSubject.teacher1 || '',
           phone1: newSubject.phone1 || '',
           teacher2: newSubject.teacher2 || '',
           phone2: newSubject.phone2 || '',
       };

      if (editingSubjectId) {
        updateSubject(editingSubjectId, subjectData);
        setEditingSubjectId(null);
      } else {
        addSubject(subjectData as Subject);
      }
      setNewSubject(filterMajorId ? { majorId: filterMajorId } : {});
    }
  };

  const handleEditSubject = (s: Subject) => {
    setNewSubject(s);
    setEditingSubjectId(s.id);
  }

  const handleCancelSubject = () => {
    setNewSubject(filterMajorId ? { majorId: filterMajorId } : {});
    setEditingSubjectId(null);
  }

  const handleSubjectFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      // Reset immediately
      e.target.value = '';

      const reader = new FileReader();
      reader.onload = (evt) => {
          try {
              const bstr = evt.target?.result;
              const wb = XLSX.read(bstr, { type: 'binary' });
              const wsname = wb.SheetNames[0];
              const ws = wb.Sheets[wsname];
              const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

              if (!data || data.length < 2) {
                  setTimeout(() => alert("File Excel rỗng hoặc không đúng định dạng!"), 100);
                  return;
              }

              const rows = data.slice(1) as any[];
              
              // Map columns: 0: Name, 1: Major Name, 2: Periods
              const newSubjects: Omit<Subject, 'id'>[] = rows.map(row => {
                  const majorName = row[1] ? String(row[1]).trim() : '';
                  const major = majors.find(m => m.name.toLowerCase() === majorName.toLowerCase());
                  
                  return {
                      name: row[0] || '',
                      majorId: major ? major.id : '', // If not found, leave empty or assign to default?
                      totalPeriods: row[2] ? Number(row[2]) : 30,
                  };
              }).filter(s => s.name);

               // Warn if some majors weren't found
               const missingMajors = newSubjects.filter(s => !s.majorId).length;
               
               if (newSubjects.length > 0) {
                   importSubjects(newSubjects);
                   let msg = `Đã nhập thành công ${newSubjects.length} môn học.`;
                   if (missingMajors > 0) {
                       msg += `\nLưu ý: Có ${missingMajors} môn chưa tìm thấy Ngành học tương ứng trong hệ thống (vui lòng kiểm tra tên Ngành trong file Excel hoặc cập nhật thủ công).`;
                   }
                   setTimeout(() => alert(msg), 100);
               } else {
                   setTimeout(() => alert("Không tìm thấy dữ liệu hợp lệ."), 100);
               }
          } catch (error) {
              console.error(error);
              setTimeout(() => alert("Lỗi khi đọc file Excel."), 100);
          }
      };
      reader.readAsBinaryString(file);
  };

  const handleClassFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      // Reset immediately
      e.target.value = '';

      const reader = new FileReader();
      reader.onload = (evt) => {
          try {
              const bstr = evt.target?.result;
              const wb = XLSX.read(bstr, { type: 'binary' });
              const wsname = wb.SheetNames[0];
              const ws = wb.Sheets[wsname];
              const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

              if (!data || data.length < 2) {
                  setTimeout(() => alert("File Excel rỗng hoặc không đúng định dạng!"), 100);
                  return;
              }

              const rows = data.slice(1) as any[];
              
              // Map columns: 0: Name, 1: Major Name, 2: Student Count, 3: School Year
              const newClasses = rows.map(row => {
                  const majorName = row[1] ? String(row[1]).trim() : '';
                  const major = majors.find(m => m.name.toLowerCase() === majorName.toLowerCase());
                  
                  return {
                      name: row[0] || '',
                      majorId: major ? major.id : '',
                      studentCount: row[2] ? Number(row[2]) : 0,
                      schoolYear: row[3] ? String(row[3]) : ''
                  };
              }).filter(c => c.name);

               const missingMajors = newClasses.filter(c => !c.majorId).length;
               
               if (newClasses.length > 0) {
                   importClasses(newClasses);
                   
                   let msg = `Đã nhập thành công ${newClasses.length} lớp.`;
                   if (missingMajors > 0) {
                       msg += `\nLưu ý: Có ${missingMajors} lớp chưa tìm thấy Ngành học tương ứng trong hệ thống.`;
                   }
                   setTimeout(() => alert(msg), 100);
               } else {
                   setTimeout(() => alert("Không tìm thấy dữ liệu hợp lệ."), 100);
               }
          } catch (error) {
              console.error(error);
              setTimeout(() => alert("Lỗi khi đọc file Excel."), 100);
          }
      };
      reader.readAsBinaryString(file);
  };

  const handleSaveClass = () => {
    if (newClass.name && newClass.majorId) {
       const classData = {
          name: newClass.name,
          studentCount: newClass.studentCount || 0,
          majorId: newClass.majorId,
          schoolYear: newClass.schoolYear || ''
      };

      if (editingClassId) {
        updateClass(editingClassId, classData);
        setEditingClassId(null);
      } else {
        addClass(classData);
      }
      setNewClass({});
    }
  };

  const handleEditClass = (c: ClassEntity) => {
    setNewClass(c);
    setEditingClassId(c.id);
  }

  const handleCancelClass = () => {
    setNewClass({});
    setEditingClassId(null);
  }

  return (
    <div className="space-y-6">
       <h1 className="text-2xl font-bold text-gray-800">Quản lý giảng dạy</h1>
       
       <div className="flex space-x-4 border-b overflow-x-auto">
         <button 
           className={`pb-2 px-4 whitespace-nowrap ${activeTab === 'teachers' ? 'border-b-2 border-blue-600 text-blue-600 font-bold' : 'text-gray-500'}`}
           onClick={() => setActiveTab('teachers')}
         >
           Giáo viên
         </button>
         <button 
           className={`pb-2 px-4 whitespace-nowrap ${activeTab === 'subjects' ? 'border-b-2 border-blue-600 text-blue-600 font-bold' : 'text-gray-500'}`}
           onClick={() => setActiveTab('subjects')}
         >
           Môn học
         </button>
         <button 
           className={`pb-2 px-4 whitespace-nowrap ${activeTab === 'classes' ? 'border-b-2 border-blue-600 text-blue-600 font-bold' : 'text-gray-500'}`}
           onClick={() => setActiveTab('classes')}
         >
           Lớp
         </button>
       </div>

       {activeTab === 'teachers' && (
         <div className="space-y-4">
            <div className="bg-white p-4 rounded shadow grid grid-cols-1 md:grid-cols-12 gap-4">
               <div className="md:col-span-2">
                   <select 
                        className="border p-2 rounded w-full" 
                        value={newTeacher.title || 'Thầy'} 
                        onChange={e => setNewTeacher({...newTeacher, title: e.target.value})}
                   >
                       <option value="Thầy">Thầy</option>
                       <option value="Cô">Cô</option>
                   </select>
               </div>
               <div className="md:col-span-4">
                   <input placeholder="Họ tên" className="border p-2 rounded w-full" value={newTeacher.name || ''} onChange={e => setNewTeacher({...newTeacher, name: e.target.value})} />
               </div>
               <div className="md:col-span-6">
                   <input placeholder="Môn dạy (Chuyên môn)" className="border p-2 rounded w-full" value={newTeacher.mainSubject || ''} onChange={e => setNewTeacher({...newTeacher, mainSubject: e.target.value})} />
               </div>

               <div className="md:col-span-3">
                   <input placeholder="Điện thoại" className="border p-2 rounded w-full" value={newTeacher.phone || ''} onChange={e => setNewTeacher({...newTeacher, phone: e.target.value})} />
               </div>
               <div className="md:col-span-2">
                   <input placeholder="Ngân hàng" className="border p-2 rounded w-full" value={newTeacher.bank || ''} onChange={e => setNewTeacher({...newTeacher, bank: e.target.value})} />
               </div>
               <div className="md:col-span-4">
                   <input placeholder="Số tài khoản" className="border p-2 rounded w-full" value={newTeacher.accountNumber || ''} onChange={e => setNewTeacher({...newTeacher, accountNumber: e.target.value})} />
               </div>
               <div className="md:col-span-3 flex gap-2">
                    <input type="number" placeholder="Thù lao/tiết" className="border p-2 rounded w-full" value={newTeacher.ratePerPeriod || ''} onChange={e => setNewTeacher({...newTeacher, ratePerPeriod: Number(e.target.value)})} />
               </div>
               
               <div className="md:col-span-12 flex flex-col lg:flex-row gap-4 justify-between items-center border-t pt-4 mt-2">
                 {/* Left: Search & Import */}
                 <div className="flex w-full lg:w-auto gap-3 items-center order-2 lg:order-1">
                    <div className="relative flex-1 lg:w-64">
                        <Search size={18} className="text-gray-400 absolute left-2 top-1/2 -translate-y-1/2" />
                        <input 
                            type="text" 
                            placeholder="Tìm kiếm giáo viên..." 
                            className="pl-9 pr-8 py-2 border rounded text-sm w-full outline-none focus:ring-2 focus:ring-blue-500"
                            value={teacherSearch}
                            onChange={(e) => setTeacherSearch(e.target.value)}
                        />
                         {teacherSearch && (
                            <button onClick={() => setTeacherSearch('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600">
                              <X size={14} />
                            </button>
                          )}
                    </div>
                    
                     <input type="file" ref={fileInputTeacherRef} accept=".xlsx,.xls" className="hidden" onChange={handleTeacherFileUpload} />
                     <button onClick={() => fileInputTeacherRef.current?.click()} className="flex items-center gap-1 text-sm bg-green-50 text-green-700 px-3 py-2 rounded hover:bg-green-100 border border-green-200 whitespace-nowrap">
                        <Upload size={16} /> <span className="hidden sm:inline">Nhập Excel</span>
                     </button>
                     <div className="relative group">
                        <HelpCircle size={18} className="text-gray-400 cursor-help" />
                        <div className="absolute left-0 bottom-full mb-2 w-64 p-2 bg-gray-800 text-white text-xs rounded shadow-lg z-10 hidden group-hover:block">
                           File Excel: <strong>Danh xưng | Họ tên | Điện thoại | Ngân hàng | Số TK | Thù lao | Môn dạy</strong>
                        </div>
                     </div>
                 </div>

                 {/* Right: Action Buttons */}
                 <div className="flex gap-2 w-full lg:w-auto justify-end order-1 lg:order-2">
                     {editingTeacherId && (
                       <button onClick={handleCancelTeacher} className="bg-gray-100 text-gray-600 px-4 py-2 rounded hover:bg-gray-200">
                         Hủy
                       </button>
                     )}
                     <button onClick={handleSaveTeacher} className={`text-white px-4 py-2 rounded flex justify-center items-center shadow-sm ${editingTeacherId ? 'bg-orange-500 hover:bg-orange-600' : 'bg-blue-600 hover:bg-blue-700'}`}>
                       {editingTeacherId ? <><Save size={18} className="mr-1"/> Cập nhật</> : <><Plus size={18} className="mr-1"/> Thêm giáo viên</>}
                     </button>
                 </div>
               </div>
            </div>

            <div className="bg-white rounded shadow overflow-hidden">
               <table className="w-full text-left">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="p-3 w-20">Danh xưng</th>
                      <th 
                        className="p-3 cursor-pointer hover:bg-gray-200 transition-colors select-none flex items-center gap-1 group"
                        onClick={toggleTeacherSort}
                        title="Bấm để sắp xếp theo tên"
                      >
                        Họ tên
                        {teacherSort === 'default' && <ArrowUpDown size={14} className="text-gray-400 group-hover:text-gray-600" />}
                        {teacherSort === 'asc' && <ArrowUp size={14} className="text-blue-600" />}
                        {teacherSort === 'desc' && <ArrowDown size={14} className="text-blue-600" />}
                      </th>
                      <th className="p-3">Môn dạy</th>
                      <th className="p-3">Điện thoại</th>
                      <th className="p-3">Ngân hàng</th>
                      <th className="p-3">Thù lao</th>
                      <th className="p-3">Hành động</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedTeachers.map(t => (
                      <tr key={t.id} className="border-t hover:bg-gray-50">
                        <td className="p-3 text-gray-600">{t.title}</td>
                        <td className="p-3 font-medium">{t.name}</td>
                        <td className="p-3 text-gray-600">{t.mainSubject}</td>
                        <td className="p-3">{t.phone}</td>
                        <td className="p-3">{t.bank} - {t.accountNumber}</td>
                        <td className="p-3">{t.ratePerPeriod.toLocaleString()}</td>
                        <td className="p-3 flex space-x-2">
                           <button onClick={() => handleEditTeacher(t)} className="text-orange-500" title="Sửa"><Edit2 size={18} /></button>
                           <button onClick={() => deleteTeacher(t.id)} className="text-red-500" title="Xóa"><Trash2 size={18} /></button>
                        </td>
                      </tr>
                    ))}
                    {sortedTeachers.length === 0 && (
                      <tr>
                        <td colSpan={7} className="p-8 text-center text-gray-400 italic">Không tìm thấy giáo viên nào.</td>
                      </tr>
                    )}
                  </tbody>
               </table>
            </div>
         </div>
       )}

       {activeTab === 'subjects' && (
         <div className="space-y-4">
             <div className="bg-white p-4 rounded shadow flex flex-col md:flex-row items-center gap-3 border border-blue-100">
                 <div className="flex items-center gap-2 w-full md:w-auto">
                    <Filter size={20} className="text-blue-600" />
                    <span className="font-semibold text-gray-700 whitespace-nowrap">Lọc theo ngành:</span>
                 </div>
                 <select 
                   className="border p-2 rounded outline-none focus:ring-2 focus:ring-blue-500 flex-1 max-w-md bg-gray-50"
                   value={filterMajorId}
                   onChange={(e) => {
                     setFilterMajorId(e.target.value);
                     setNewSubject(prev => ({ ...prev, majorId: e.target.value }));
                   }}
                 >
                    <option value="">-- Tất cả các ngành --</option>
                    {majors.map(m => <option key={m.id} value={m.id}>{m.name}</option>)}
                 </select>

                 <div className="flex items-center gap-2 w-full md:w-auto flex-1">
                    <Search size={20} className="text-gray-400" />
                    <input 
                        type="text" 
                        placeholder="Tìm tên môn học..." 
                        className="border p-2 rounded outline-none text-sm w-full"
                        value={subjectSearch}
                        onChange={(e) => setSubjectSearch(e.target.value)}
                    />
                     {subjectSearch && (
                        <button onClick={() => setSubjectSearch('')} className="text-gray-400 hover:text-gray-600">
                          <X size={16} />
                        </button>
                      )}
                 </div>

                 <div className="flex items-center gap-2 ml-auto">
                     <input type="file" ref={fileInputSubjectRef} accept=".xlsx,.xls" className="hidden" onChange={handleSubjectFileUpload} />
                     <button onClick={() => fileInputSubjectRef.current?.click()} className="flex items-center gap-1 text-sm bg-green-50 text-green-700 px-3 py-1.5 rounded hover:bg-green-100 border border-green-200">
                        <Upload size={16} /> Nhập Excel
                     </button>
                      <div className="relative group">
                        <HelpCircle size={18} className="text-gray-400 cursor-help" />
                        <div className="absolute right-0 top-8 w-64 p-2 bg-gray-800 text-white text-xs rounded shadow-lg z-10 hidden group-hover:block">
                           File Excel: <strong>Tên môn | Ngành học | Số tiết</strong>
                           <br/><span className="text-[10px] opacity-75">Tên Ngành cần khớp với hệ thống.</span>
                        </div>
                     </div>
                 </div>
             </div>

             <div className="bg-white p-4 rounded shadow space-y-4">
               <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-end">
                  <div className="md:col-span-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">Tên môn học</label>
                    <input className="border p-2 rounded w-full" value={newSubject.name || ''} onChange={e => setNewSubject({...newSubject, name: e.target.value})} placeholder="Nhập tên môn..." />
                  </div>
                  <div className="md:col-span-3">
                    <label className="block text-sm font-medium text-gray-700 mb-1">Ngành</label>
                    <select 
                        className="border p-2 rounded w-full" 
                        value={newSubject.majorId || filterMajorId || ''} 
                        onChange={e => setNewSubject({...newSubject, majorId: e.target.value})}
                    >
                        <option value="">Chọn Ngành...</option>
                        {majors.map(m => <option key={m.id} value={m.id}>{m.name}</option>)}
                    </select>
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-medium text-gray-700 mb-1">Tổng số tiết</label>
                    <input type="number" className="border p-2 rounded w-full" value={newSubject.totalPeriods || ''} onChange={e => setNewSubject({...newSubject, totalPeriods: Number(e.target.value)})} placeholder="0" />
                  </div>
                   <div className="md:col-span-3 pb-2">
                    <label className="flex items-center space-x-2 cursor-pointer select-none">
                        <input 
                            type="checkbox" 
                            className="w-5 h-5 text-blue-600 rounded" 
                            checked={newSubject.isShared || false}
                            onChange={e => setNewSubject({...newSubject, isShared: e.target.checked})}
                        />
                        <span className="text-gray-700 font-medium flex items-center">
                            Môn chung (Lớp ghép)
                            <Users size={14} className="ml-1 text-gray-400" />
                        </span>
                    </label>
                  </div>
               </div>

               <div className="grid grid-cols-1 md:grid-cols-2 gap-4 border-t pt-4 items-end">
                  <div className="space-y-1">
                     <p className="text-sm font-bold text-gray-600">Giáo viên phụ trách 1</p>
                     <select
                        className="border p-2 rounded w-full text-sm"
                        value={newSubject.teacher1 || ''}
                        onChange={(e) => {
                            const selectedName = e.target.value;
                            const teacher = teachers.find(t => t.name === selectedName);
                            setNewSubject({
                                ...newSubject,
                                teacher1: selectedName,
                                phone1: teacher ? teacher.phone : ''
                            });
                        }}
                     >
                        <option value="">-- Chọn giáo viên --</option>
                        {teachers.map(t => (
                            <option key={t.id} value={t.name}>{t.name}</option>
                        ))}
                     </select>
                  </div>
                  <div className="space-y-1">
                     <p className="text-sm font-bold text-gray-600">Giáo viên phụ trách 2</p>
                     <div className="flex items-center gap-2">
                        <select
                            className="border p-2 rounded w-full text-sm"
                            value={newSubject.teacher2 || ''}
                            onChange={(e) => {
                                const selectedName = e.target.value;
                                const teacher = teachers.find(t => t.name === selectedName);
                                setNewSubject({
                                    ...newSubject,
                                    teacher2: selectedName,
                                    phone2: teacher ? teacher.phone : ''
                                });
                            }}
                        >
                            <option value="">-- Chọn giáo viên --</option>
                            {teachers.map(t => (
                                <option key={t.id} value={t.name}>{t.name}</option>
                            ))}
                        </select>
                        <div className="flex gap-2">
                             {editingSubjectId && (
                               <button onClick={handleCancelSubject} className="bg-gray-300 text-gray-700 p-2 rounded px-3 text-sm whitespace-nowrap h-[38px]">Hủy</button>
                             )}
                             <button onClick={handleSaveSubject} className={`text-white p-2 rounded px-4 flex items-center text-sm whitespace-nowrap h-[38px] ${editingSubjectId ? 'bg-orange-500' : 'bg-green-600'}`}>
                               {editingSubjectId ? <><Save size={16} className="mr-1"/> Cập nhật</> : <><Plus size={16} className="mr-1"/> Thêm</>}
                             </button>
                         </div>
                     </div>
                  </div>
               </div>
            </div>

             <div className="bg-white rounded shadow overflow-hidden">
               <table className="w-full text-left">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="p-3">Tên môn</th>
                      <th className="p-3">Ngành</th>
                      <th className="p-3">Tổng tiết</th>
                      <th className="p-3">GV Phụ trách</th>
                      <th className="p-3">Hành động</th>
                    </tr>
                  </thead>
                  <tbody>
                    {subjects
                      .filter(s => (!filterMajorId || s.majorId === filterMajorId) && s.name.toLowerCase().includes(subjectSearch.toLowerCase()))
                      .map(s => (
                      <tr key={s.id} className="border-t hover:bg-gray-50">
                        <td className="p-3 align-top">
                            <div className="font-medium">{s.name}</div>
                            {s.isShared && (
                                <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-purple-100 text-purple-800 mt-1">
                                    <Users size={10} className="mr-1" /> Môn chung
                                </span>
                            )}
                        </td>
                        <td className="p-3 align-top">{majors.find(m => m.id === s.majorId)?.name}</td>
                        <td className="p-3 align-top">{s.totalPeriods}</td>
                        <td className="p-3 text-sm text-gray-600 align-top">
                           {s.teacher1 && <div>1. {s.teacher1} - {s.phone1}</div>}
                           {s.teacher2 && <div>2. {s.teacher2} - {s.phone2}</div>}
                           {!s.teacher1 && !s.teacher2 && <span className="italic text-gray-400">Chưa cập nhật</span>}
                        </td>
                        <td className="p-3 flex space-x-2 align-top">
                           <button onClick={() => handleEditSubject(s)} className="text-orange-500" title="Sửa"><Edit2 size={18} /></button>
                           <button onClick={() => deleteSubject(s.id)} className="text-red-500" title="Xóa"><Trash2 size={18} /></button>
                        </td>
                      </tr>
                    ))}
                    {subjects.filter(s => (!filterMajorId || s.majorId === filterMajorId) && s.name.toLowerCase().includes(subjectSearch.toLowerCase())).length === 0 && (
                      <tr>
                        <td colSpan={5} className="p-8 text-center text-gray-400 italic">
                            {subjectSearch 
                             ? `Không tìm thấy môn nào có tên "${subjectSearch}" ${filterMajorId ? 'trong ngành này' : ''}.` 
                             : `Không có môn học nào ${filterMajorId ? 'trong ngành này' : ''}.`}
                        </td>
                      </tr>
                    )}
                  </tbody>
               </table>
            </div>
         </div>
       )}

       {activeTab === 'classes' && (
         <div className="space-y-4">
             <div className="bg-white p-4 rounded shadow grid grid-cols-1 md:grid-cols-5 gap-4">
               <input placeholder="Tên lớp (25DC...)" className="border p-2 rounded md:col-span-2" value={newClass.name || ''} onChange={e => setNewClass({...newClass, name: e.target.value})} />
               <select className="border p-2 rounded" value={newClass.majorId || ''} onChange={e => setNewClass({...newClass, majorId: e.target.value})}>
                  <option value="">Chọn Ngành...</option>
                  {majors.map(m => <option key={m.id} value={m.id}>{m.name}</option>)}
               </select>
               <div className="flex gap-2">
                 <input type="number" placeholder="Số lượng" className="border p-2 rounded w-1/2" value={newClass.studentCount || ''} onChange={e => setNewClass({...newClass, studentCount: Number(e.target.value)})} />
                 <input placeholder="Niên khóa" className="border p-2 rounded w-1/2" value={newClass.schoolYear || ''} onChange={e => setNewClass({...newClass, schoolYear: e.target.value})} />
               </div>
               
               <div className="flex gap-2">
                 <button onClick={handleSaveClass} className={`flex-1 text-white p-2 rounded flex justify-center items-center ${editingClassId ? 'bg-orange-500' : 'bg-purple-600'}`}>
                   {editingClassId ? <><Save size={18} className="mr-1"/> Cập nhật</> : <><Plus size={18} className="mr-1"/> Thêm</>}
                 </button>
                 {editingClassId && (
                   <button onClick={handleCancelClass} className="bg-gray-300 text-gray-700 p-2 rounded">
                     <X size={18} />
                   </button>
                 )}
               </div>
            </div>

             <div className="bg-white p-3 rounded shadow border border-gray-100 flex flex-col md:flex-row items-center gap-3">
              <div className="flex items-center gap-2 flex-1 w-full">
                  <Search size={20} className="text-gray-400" />
                  <input 
                    type="text" 
                    placeholder="Tìm kiếm lớp..." 
                    className="flex-1 outline-none text-sm"
                    value={classSearch}
                    onChange={(e) => setClassSearch(e.target.value)}
                  />
                  {classSearch && (
                    <button onClick={() => setClassSearch('')} className="text-gray-400 hover:text-gray-600">
                      <X size={16} />
                    </button>
                  )}
              </div>
              <div className="flex items-center gap-2">
                 <input type="file" ref={fileInputClassRef} accept=".xlsx,.xls" className="hidden" onChange={handleClassFileUpload} />
                 <button onClick={() => fileInputClassRef.current?.click()} className="flex items-center gap-1 text-sm bg-green-50 text-green-700 px-3 py-1.5 rounded hover:bg-green-100 border border-green-200">
                    <Upload size={16} /> Nhập Excel
                 </button>
                 <div className="relative group">
                    <HelpCircle size={18} className="text-gray-400 cursor-help" />
                    <div className="absolute right-0 top-8 w-64 p-2 bg-gray-800 text-white text-xs rounded shadow-lg z-10 hidden group-hover:block">
                       File Excel: <strong>Tên lớp | Ngành học | Số lượng | Niên khóa</strong>
                       <br/><span className="text-[10px] opacity-75">Tên Ngành cần khớp với hệ thống.</span>
                    </div>
                 </div>
              </div>
            </div>

             <div className="bg-white rounded shadow overflow-hidden">
               <table className="w-full text-left">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="p-3">Tên lớp</th>
                      <th 
                        className="p-3 cursor-pointer hover:bg-gray-200 transition-colors select-none flex items-center gap-1 group"
                        onClick={toggleClassSortMajor}
                        title="Bấm để sắp xếp theo ngành"
                      >
                        Ngành
                        {classSortMajor === 'default' && <ArrowUpDown size={14} className="text-gray-400 group-hover:text-gray-600" />}
                        {classSortMajor === 'asc' && <ArrowUp size={14} className="text-blue-600" />}
                        {classSortMajor === 'desc' && <ArrowDown size={14} className="text-blue-600" />}
                      </th>
                      <th className="p-3">Số lượng</th>
                      <th className="p-3">Niên khóa</th>
                      <th className="p-3">Hành động</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedClasses.map(c => (
                      <tr key={c.id} className="border-t hover:bg-gray-50">
                        <td className="p-3 font-medium">{c.name}</td>
                        <td className="p-3">{majors.find(m => m.id === c.majorId)?.name}</td>
                        <td className="p-3">{c.studentCount}</td>
                        <td className="p-3">{c.schoolYear}</td>
                        <td className="p-3 flex space-x-2">
                           <button onClick={() => handleEditClass(c)} className="text-orange-500" title="Sửa"><Edit2 size={18} /></button>
                           <button onClick={() => deleteClass(c.id)} className="text-red-500" title="Xóa"><Trash2 size={18} /></button>
                        </td>
                      </tr>
                    ))}
                    {sortedClasses.length === 0 && (
                      <tr>
                        <td colSpan={5} className="p-8 text-center text-gray-400 italic">
                            {classSearch 
                             ? `Không tìm thấy lớp học nào phù hợp với "${classSearch}".` 
                             : `Không tìm thấy lớp học nào.`}
                        </td>
                      </tr>
                    )}
                  </tbody>
               </table>
            </div>
         </div>
       )}
    </div>
  );
};

export default Management;