import React, { useMemo, useState } from 'react';
import { useApp } from '../store/AppContext';
import { ScheduleStatus } from '../types';
import XLSX from 'xlsx';
import ExcelJS from 'exceljs'; // Import ExcelJS
import saveAs from 'file-saver';
import { Download, Trash2, CheckCircle, CreditCard, FileSpreadsheet } from 'lucide-react';
import { format } from 'date-fns';
import { parseLocal, base64ToArrayBuffer } from '../utils';

const Payment: React.FC = () => {
  const { teachers, schedules, subjects, classes, templates } = useApp();

  // State to track paid/removed subjects (persisted in localStorage)
  // Store format: "subjectId-classId"
  const [paidItems, setPaidItems] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('paid_completed_subjects');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  // Calculate Completed Subjects (Finished)
  const completedSubjects = useMemo(() => {
    const results: any[] = [];
    
    classes.forEach(cls => {
        const classSubjects = subjects.filter(s => s.majorId === cls.majorId);
        
        classSubjects.forEach(sub => {
            // Unique key for payment tracking
            const uniqueKey = `${sub.id}-${cls.id}`;

            // Skip if already paid/deleted
            if (paidItems.includes(uniqueKey)) return;

             const relevantSchedules = schedules.filter(sch => 
                sch.subjectId === sub.id && 
                sch.classId === cls.id && 
                sch.status !== ScheduleStatus.OFF
            );
            
            const learned = relevantSchedules.reduce((acc, curr) => acc + curr.periodCount, 0);
            
            // Condition: Learned >= Total AND started (learned > 0)
            if (learned >= sub.totalPeriods && learned > 0) {
                 // Get list of teachers who taught this subject for this class
                 const teacherIds = Array.from(new Set(relevantSchedules.map(s => s.teacherId)));
                 const teacherNames = teacherIds.map(tid => teachers.find(t => t.id === tid)?.name || 'GV đã xóa').join(', ');

                 results.push({
                     subjectId: sub.id,
                     classId: cls.id,
                     uniqueKey: uniqueKey,
                     subjectName: sub.name,
                     className: cls.name,
                     teacherName: teacherNames || "Chưa xác định",
                     totalPeriods: sub.totalPeriods
                 });
            }
        });
    });
    
    return results;
  }, [subjects, schedules, classes, teachers, paidItems]);

  const handleDelete = (key: string) => {
    setPaidItems(prev => {
        const newPaidItems = [...prev, key];
        localStorage.setItem('paid_completed_subjects', JSON.stringify(newPaidItems));
        return newPaidItems;
    });
  };

  const getFilteredSchedules = (item: any) => {
      return schedules
      .filter(s => {
          if (s.subjectId !== item.subjectId || s.classId !== item.classId || s.status === ScheduleStatus.OFF) return false;
          
          if (s.type === 'class') return true;
          if (s.type === 'exam') {
              return s.note && s.note.toLowerCase().includes('thực hành');
          }
          return false;
      })
      .sort((a, b) => {
          // Sort by date then startPeriod
          const da = new Date(a.date).getTime();
          const db = new Date(b.date).getTime();
          return da - db || a.startPeriod - b.startPeriod;
      });
  }

  const exportFinishedSubject = (item: any) => {
    const relevantSchedules = getFilteredSchedules(item);

    const data = relevantSchedules.map(s => ({
       'Giáo viên giảng dạy': teachers.find(t => t.id === s.teacherId)?.name || 'GV đã xóa',
       'Ngày dạy': format(parseLocal(s.date), 'dd/MM/yyyy'),
       'Số tiết dạy': s.periodCount,
       'Lớp': item.className,
       'Loại': s.type === 'exam' ? 'Thi' : 'Học',
       'Ghi chú': s.note || ''
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{ wch: 25 }, { wch: 15 }, { wch: 12 }, { wch: 20 }, { wch: 10 }, { wch: 30 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ChiTietMonHoc");
    XLSX.writeFile(wb, `ThongKe_${item.subjectName}_${item.className}.xlsx`);
  };

  const exportExcelTemplate = async (item: any) => {
      // 1. Check if templates exist
      const excelTemplates = templates.filter(t => t.type === 'payment_excel');
      if (excelTemplates.length === 0) {
          alert('Chưa có mẫu Excel nào. Vui lòng vào mục "Hệ thống" để tải lên mẫu .xlsx!');
          return;
      }

      // 2. Select Template (Use first one by default)
      const template = excelTemplates[0];

      try {
          // 3. Prepare Data
          const relevantSchedules = getFilteredSchedules(item);
          const firstSchedule = relevantSchedules[0];
          const teacherObj = firstSchedule ? teachers.find(t => t.id === firstSchedule.teacherId) : null;
          const startDate = relevantSchedules.length > 0 ? format(parseLocal(relevantSchedules[0].date), 'dd/MM/yyyy') : '...';
          const endDate = relevantSchedules.length > 0 ? format(parseLocal(relevantSchedules[relevantSchedules.length - 1].date), 'dd/MM/yyyy') : '...';
          const actualTotalPeriods = relevantSchedules.reduce((sum, s) => sum + s.periodCount, 0);

          const replacements: Record<string, string> = {
              "{teacherTitle}": teacherObj?.title || 'Thầy/Cô', 
              "{teacherName}": item.teacherName,
              "{teacherPhone}": teacherObj?.phone || '',
              "{teacherBank}": teacherObj?.bank || '',
              "{teacherAccount}": teacherObj?.accountNumber || '',
              "{subjectName}": item.subjectName,
              "{className}": item.className, // Scalar replacement (Header)
              "{totalPeriods}": String(item.totalPeriods),
              "{actualTotalPeriods}": String(actualTotalPeriods),
              "{fromDate}": startDate,
              "{toDate}": endDate,
          };

          // 4. Load Template using ExcelJS
          const buffer = base64ToArrayBuffer(template.content);
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(buffer);
          
          const ws = workbook.worksheets[0];
          if (!ws) {
              alert("File mẫu không hợp lệ.");
              return;
          }

          // 5. Detect Scalar Replacements & List Structure (Multi-column support)
          let listRowIndex = -1;
          // colMap structure: { 1: { date: col, periods: col }, 2: { date: col, periods: col } }
          const colSets: Record<number, Record<string, number>> = {};

          ws.eachRow((row, rowNumber) => {
              row.eachCell((cell, colNumber) => {
                  let val = cell.value ? String(cell.value) : '';

                  // Detect list placeholders with optional suffix (e.g., {date}, {date_2})
                  // Regex: matches {date}, {date_1}, {date_2}, etc.
                  const listVars = ['date', 'periods', 'type', 'note', 'className'];
                  
                  for (const v of listVars) {
                      const regex = new RegExp(`{${v}(_([0-9]+))?}`);
                      const match = val.match(regex);
                      if (match) {
                          listRowIndex = rowNumber;
                          const suffix = match[2] ? parseInt(match[2]) : 1;
                          
                          if (!colSets[suffix]) colSets[suffix] = {};
                          colSets[suffix][v] = colNumber;
                      }
                  }

                  // Replace Scalars (Headers)
                  Object.keys(replacements).forEach(key => {
                      if(val.includes(key)) {
                          val = val.replace(key, replacements[key]);
                          cell.value = val;
                      }
                  });
              });
          });

          // 6. Handle Multi-Column List Insertion
          const setKeys = Object.keys(colSets).map(Number).sort((a,b) => a-b);
          const numSets = setKeys.length;

          // Helper to clean placeholders in a row (if no data)
          const cleanRowPlaceholders = (row: any) => {
             row.eachCell((cell: any) => {
                  const val = String(cell.value);
                  if (val.match(/{date(_\d+)?}/) || val.match(/{periods(_\d+)?}/) || val.match(/{className(_\d+)?}/)) {
                      cell.value = null;
                  }
             });
          };

          if (listRowIndex !== -1 && relevantSchedules.length > 0 && numSets > 0) {
              
              // Determine how many items per set (Divide evenly)
              const itemsPerSet = Math.ceil(relevantSchedules.length / numSets);
              
              // Helper to fill data into a specific row for a specific set
              const fillDataToCell = (row: any, schedule: any, setIndex: number) => {
                  const map = colSets[setIndex];
                  if (!map) return;
                  
                  if(map['date']) row.getCell(map['date']).value = format(parseLocal(schedule.date), 'dd/MM/yyyy');
                  if(map['periods']) row.getCell(map['periods']).value = schedule.periodCount;
                  if(map['type']) row.getCell(map['type']).value = schedule.type === 'exam' ? 'Thi' : 'Học';
                  if(map['note']) row.getCell(map['note']).value = schedule.note || '';
                  if(map['className']) row.getCell(map['className']).value = item.className; // From item context

                  // Clean up placeholders in this cell if they remain
                  // Note: Logic here assumes the variable took up the whole cell or is ready to be overwritten
              };

              // Process:
              // We need to iterate 'itemsPerSet' times.
              // For each iteration i, we fill Set 1 with data[i], Set 2 with data[i + itemsPerSet], etc.
              
              // 1. Fill the FIRST row (existing in template)
              for (let setIdx = 1; setIdx <= numSets; setIdx++) {
                  const dataIndex = (setIdx - 1) * itemsPerSet + 0; // i=0
                  if (dataIndex < relevantSchedules.length) {
                      fillDataToCell(ws.getRow(listRowIndex), relevantSchedules[dataIndex], setIdx);
                  }
              }

              // 2. Insert and Fill remaining rows
              for (let i = 1; i < itemsPerSet; i++) {
                  const newRowIdx = listRowIndex + i;
                  const newRow = ws.insertRow(newRowIdx, [], 'i'); // Inherit style
                  
                  for (let setIdx = 1; setIdx <= numSets; setIdx++) {
                      const dataIndex = (setIdx - 1) * itemsPerSet + i;
                      if (dataIndex < relevantSchedules.length) {
                          fillDataToCell(newRow, relevantSchedules[dataIndex], setIdx);
                      }
                  }
              }
              
              // Scan the used rows to ensure no placeholders remain (for short columns)
               for (let i = 0; i < itemsPerSet; i++) {
                   const rowToCheck = ws.getRow(listRowIndex + i);
                   cleanRowPlaceholders(rowToCheck);
               }

          } else if (listRowIndex !== -1) {
              // Found placeholder row but no data -> Clear placeholders
              const row = ws.getRow(listRowIndex);
              cleanRowPlaceholders(row);
          }

          // 7. Write Buffer and Save
          const outBuffer = await workbook.xlsx.writeBuffer();
          const blob = new Blob([outBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
          saveAs(blob, `PhieuThanhToan_${item.subjectName}.xlsx`);

      } catch (error) {
          console.error(error);
          alert("Lỗi khi xuất file Excel: " + error);
      }
  };

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold text-gray-800 flex items-center">
        <CreditCard className="mr-3 text-blue-600" /> Thanh toán giảng dạy
      </h1>
      
      <div className="bg-white p-6 rounded-xl shadow border border-gray-200">
        <h3 className="font-bold text-gray-800 text-lg mb-4 flex items-center">
            <CheckCircle className="mr-2 text-green-500" size={20} />
            Môn học đã kết thúc (Chờ thanh toán)
        </h3>
        
        {completedSubjects.length > 0 ? (
            <div className="border border-gray-200 rounded-lg overflow-hidden">
                <table className="w-full text-left bg-white">
                    <thead className="bg-gray-50 border-b border-gray-200">
                        <tr>
                            <th className="p-4 text-gray-600 font-semibold text-sm">Thông tin môn học</th>
                            <th className="p-4 text-gray-600 font-semibold text-sm w-48 text-center">Thao tác</th>
                            <th className="p-4 text-gray-600 font-semibold text-sm w-32 text-center">Trạng thái</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                        {completedSubjects.map((item: any) => (
                            <tr key={item.uniqueKey} className="hover:bg-gray-50 transition-colors">
                                <td className="p-4 text-gray-800 font-medium text-[15px]">
                                    {item.subjectName} - GV: {item.teacherName} - Lớp: {item.className}
                                </td>
                                <td className="p-4 text-center">
                                    <div className="flex justify-center gap-2">
                                        <button 
                                            onClick={() => exportFinishedSubject(item)}
                                            className="bg-green-100 hover:bg-green-200 text-green-800 px-3 py-1.5 rounded text-sm font-medium whitespace-nowrap transition-colors flex items-center"
                                            title="Xuất bảng kê chi tiết (Excel)"
                                        >
                                            <Download size={14} className="mr-1" /> Excel
                                        </button>
                                        <button 
                                            onClick={() => exportExcelTemplate(item)}
                                            className="bg-orange-100 hover:bg-orange-200 text-orange-800 px-3 py-1.5 rounded text-sm font-medium whitespace-nowrap transition-colors flex items-center"
                                            title="Xuất phiếu thanh toán theo mẫu (Excel)"
                                        >
                                            <FileSpreadsheet size={14} className="mr-1" /> Excel (Mẫu)
                                        </button>
                                    </div>
                                </td>
                                <td className="p-4 text-center">
                                    <button
                                        onClick={() => handleDelete(item.uniqueKey)}
                                        className="bg-red-50 hover:bg-red-100 text-red-600 px-3 py-1.5 rounded text-sm font-medium transition-colors flex items-center justify-center mx-auto"
                                    >
                                        <Trash2 size={14} className="mr-1" /> Xóa
                                    </button>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        ) : (
            <div className="text-gray-500 italic py-12 text-center bg-gray-50 rounded border border-dashed border-gray-200">
                Chưa có môn học nào hoàn thành (hoặc tất cả đã được thanh toán).
            </div>
        )}
      </div>
    </div>
  );
};

export default Payment;