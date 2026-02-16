import React, { useRef } from 'react';
import { useApp } from '../store/AppContext';
import { AppState } from '../types';
import { Download, Upload, RefreshCcw, FileJson, Settings, FileSpreadsheet, Trash2, FileQuestion, FileText } from 'lucide-react';
import { format } from 'date-fns';

const SystemManager: React.FC = () => {
  const { 
    teachers, subjects, majors, classes, schedules, students, documents, templates, holidays,
    loadData, resetData, addTemplate, deleteTemplate
  } = useApp();
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const templatePaymentInputRef = useRef<HTMLInputElement>(null);
  const templateStudentInputRef = useRef<HTMLInputElement>(null);
  const templateInvitationInputRef = useRef<HTMLInputElement>(null);

  // Backup Handler
  const handleBackup = () => {
      const data: AppState = { teachers, subjects, majors, classes, schedules, students, documents, templates, holidays };
      const json = JSON.stringify(data, null, 2);
      const blob = new Blob([json], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `EduPro_Backup_${format(new Date(), 'dd-MM-yyyy')}.json`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
  };

  const handleRestoreClick = () => {
      fileInputRef.current?.click();
  };

  const handleRestoreFile = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (evt) => {
          try {
              const json = evt.target?.result as string;
              const data = JSON.parse(json) as AppState;
              loadData(data);
              alert("Khôi phục dữ liệu thành công!");
          } catch (error) {
              alert("Lỗi: File dữ liệu không hợp lệ.");
          }
          if (fileInputRef.current) fileInputRef.current.value = '';
      };
      reader.readAsText(file);
  };

  const handleResetWithBackup = () => {
      const confirmMsg = "CẢNH BÁO QUAN TRỌNG:\n\nBạn đang thực hiện xóa toàn bộ dữ liệu hệ thống (bao gồm cả dữ liệu mẫu). Hành động này không thể hoàn tác.\n\nĐể đảm bảo an toàn, hệ thống sẽ tự động lưu một bản dự phòng dữ liệu hiện tại trước khi xóa.\n\nBạn có chắc chắn muốn tiếp tục?";
      if (window.confirm(confirmMsg)) {
          // 1. Auto Backup
          handleBackup();
          
          // 2. Reset with delay to ensure download starts
          setTimeout(() => {
              resetData();
              alert("Đã sao lưu và xóa toàn bộ dữ liệu hệ thống.");
          }, 500);
      }
  };

  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'payment_excel' | 'student_list_excel' | 'invitation_word') => {
    const file = e.target.files?.[0];
    if (!file) return;

    const isExcel = type.includes('excel');
    const requiredExt = isExcel ? '.xlsx' : '.docx';

    if (!file.name.endsWith(requiredExt)) {
        alert(`Vui lòng chọn file ${isExcel ? 'Excel (.xlsx)' : 'Word (.docx)'}`);
        return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
        const content = evt.target?.result as string;
        addTemplate({
            name: file.name,
            type: type,
            content: content
        });
        
        // Reset inputs
        if (templatePaymentInputRef.current) templatePaymentInputRef.current.value = '';
        if (templateStudentInputRef.current) templateStudentInputRef.current.value = '';
        if (templateInvitationInputRef.current) templateInvitationInputRef.current.value = '';
        
        alert(`Đã thêm mẫu thành công!`);
    };
    reader.readAsDataURL(file);
  };

  // Generate and download instruction file
  const handleDownloadGuide = (type: 'payment' | 'student' | 'invitation') => {
      let content = "";
      let filename = "";

      if (type === 'payment') {
          filename = "Huong_dan_bien_Excel_Thanh_toan.txt";
          content = `HƯỚNG DẪN CÁC BIẾN MẪU EXCEL (THANH TOÁN GIẢNG DẠY)
-------------------------------------------------------
Cách dùng: Nhập tên biến (bao gồm cả dấu ngoặc nhọn) vào ô Excel bất kỳ. Hệ thống sẽ tự động thay thế bằng dữ liệu thực tế.

1. THÔNG TIN CHUNG (Thay thế giá trị đơn):
   {teacherName}        : Tên giáo viên
   {teacherTitle}       : Danh xưng (Thầy/Cô)
   {teacherPhone}       : Số điện thoại
   {teacherBank}        : Tên ngân hàng
   {teacherAccount}     : Số tài khoản
   {subjectName}        : Tên môn học
   {className}          : Tên lớp
   {totalPeriods}       : Tổng số tiết theo chương trình
   {actualTotalPeriods} : Tổng số tiết thực tế đã dạy
   {fromDate}           : Ngày bắt đầu dạy
   {toDate}             : Ngày kết thúc dạy

2. BẢNG KÊ CHI TIẾT (Danh sách các buổi học):
   Để tạo bảng kê, hãy đặt các biến sau vào MỘT HÀNG. Hệ thống sẽ tự động lặp lại hàng này cho mỗi buổi dạy.
   
   {date}               : Ngày dạy (dd/MM/yyyy)
   {periods}            : Số tiết dạy của buổi đó
   {type}               : Nội dung (Học / Thi)
   {note}               : Ghi chú
`;
      } else if (type === 'student') {
          filename = "Huong_dan_bien_Excel_Danh_sach_lop.txt";
          content = `HƯỚNG DẪN CÁC BIẾN MẪU EXCEL (DANH SÁCH LỚP)
-------------------------------------------------------
Cách dùng: Nhập tên biến (bao gồm cả dấu ngoặc nhọn) vào ô Excel bất kỳ.

1. THÔNG TIN LỚP HỌC:
   {className}          : Tên lớp
   {majorName}          : Tên ngành/nghề
   {studentCount}       : Sĩ số
   {schoolYear}         : Niên khóa

2. DANH SÁCH SINH VIÊN (Đặt trên 1 hàng để tạo bảng):
   {stt}                : Số thứ tự
   {studentCode}        : Mã số sinh viên
   {firstName}          : Tên
   {lastName}           : Họ và tên đệm
   {fullName}           : Họ và tên đầy đủ
   {dob}                : Ngày sinh
   {pob}                : Nơi sinh
   {phone}              : Số điện thoại
`;
      } else {
          filename = "Huong_dan_bien_Word_Thu_moi.txt";
          content = `HƯỚNG DẪN CÁC BIẾN MẪU WORD (THƯ MỜI GIẢNG)
-------------------------------------------------------
Cách dùng: Copy các biến bên dưới (bao gồm cả dấu ngoặc nhọn) và dán vào file Word mẫu (.docx).

1. THÔNG TIN CÁ NHÂN & MÔN HỌC:
   {teacherName}        : Tên giáo viên
   {teacherTitle}       : Danh xưng (Thầy/Cô)
   {subjectName}        : Tên môn học
   {className}          : Lớp dạy
   {totalPeriods}       : Tổng số tiết
   {rate}               : Thù lao tiết dạy (đã định dạng tiền tệ)

2. THỜI GIAN & ĐỊA ĐIỂM:
   {dates}              : Thời gian dạy (VD: Từ ngày 01/01/2024 đến ngày 30/01/2024)
   {sessions}           : Tiết dạy (VD: Sáng (1-5), Chiều (6-10))
   {room}               : Phòng học (liệt kê các phòng)
   
3. CƠ SỞ ĐÀO TẠO (Tự động xác định dựa trên tên lớp):
   {location}           : Địa chỉ cơ sở (Cơ sở 1 hoặc Cơ sở 2)
   {mapLink}            : Link Google Maps tương ứng
   
   Quy tắc:
   - Nếu tên lớp kết thúc bằng '1' -> Cơ sở 1 (Bình Hoà, TP.HCM)
   - Nếu tên lớp kết thúc bằng '2' -> Cơ sở 2 (Tân Đông Hiệp, TP.HCM)
`;
      }

      const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
  };

  return (
    <div className="space-y-6">
        <h1 className="text-2xl font-bold text-gray-800">Hệ thống & Cấu hình</h1>

        {/* Backup / Restore Section */}
        <div className="bg-white p-6 rounded shadow border border-blue-100">
            <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
                <Settings className="mr-2 text-blue-600" /> Sao lưu và Khôi phục dữ liệu
            </h2>
             {/* ... (Giữ nguyên nội dung Backup/Restore) ... */}
             <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="border border-green-200 bg-green-50 p-6 rounded-lg">
                    <h3 className="font-bold text-green-800 mb-2 flex items-center">
                            <Download className="mr-2" size={20}/> 1. Sao lưu dữ liệu (Backup)
                    </h3>
                    <p className="text-sm text-green-700 mb-4">
                        Tải toàn bộ dữ liệu hiện tại (Giáo viên, Lịch, Học sinh...) về máy tính dưới dạng file .JSON.
                    </p>
                    <button 
                        onClick={handleBackup}
                        className="bg-green-600 text-white px-4 py-2 rounded shadow hover:bg-green-700 flex items-center"
                    >
                        <FileJson className="mr-2" size={18} /> Tải file sao lưu
                    </button>
                </div>

                <div className="border border-blue-200 bg-blue-50 p-6 rounded-lg">
                    <h3 className="font-bold text-blue-800 mb-2 flex items-center">
                            <Upload className="mr-2" size={20}/> 2. Khôi phục dữ liệu (Restore)
                    </h3>
                    <p className="text-sm text-blue-700 mb-4">
                        Chọn file .JSON đã sao lưu để khôi phục lại dữ liệu. <br/>
                        <span className="font-bold text-red-500">Lưu ý: Dữ liệu hiện tại sẽ bị ghi đè hoàn toàn.</span>
                    </p>
                    <input 
                        type="file" 
                        accept=".json" 
                        ref={fileInputRef} 
                        className="hidden" 
                        onChange={handleRestoreFile} 
                    />
                    <button 
                        onClick={handleRestoreClick}
                        className="bg-blue-600 text-white px-4 py-2 rounded shadow hover:bg-blue-700 flex items-center"
                    >
                        <Upload className="mr-2" size={18} /> Chọn file để khôi phục
                    </button>
                </div>
            </div>

            <div className="mt-8 pt-6 border-t">
                <h3 className="font-bold text-red-600 mb-2 flex items-center">
                    <RefreshCcw className="mr-2" size={18} /> Vùng nguy hiểm
                </h3>
                <button 
                    onClick={handleResetWithBackup}
                    className="text-red-500 border border-red-200 hover:bg-red-50 px-4 py-2 rounded transition"
                >
                    Xóa toàn bộ dữ liệu & Reset về mặc định
                </button>
            </div>
        </div>

        {/* Invitation Template Management (New Section) */}
        <div className="bg-white p-6 rounded shadow border border-gray-200">
             <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
                <FileText className="mr-2 text-purple-600" /> Quản lý mẫu Thư mời giảng (Word)
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                     <p className="text-sm text-gray-600 mb-3">
                         Tải lên các file Word (.docx) mẫu để sử dụng cho chức năng xuất thư mời giảng.
                     </p>
                     <input 
                        type="file" 
                        accept=".docx" 
                        ref={templateInvitationInputRef} 
                        className="hidden" 
                        onChange={(e) => handleTemplateUpload(e, 'invitation_word')} 
                    />
                    <div className="flex flex-wrap gap-2 mb-4">
                        <button 
                            onClick={() => templateInvitationInputRef.current?.click()}
                            className="bg-purple-600 text-white px-4 py-2 rounded shadow hover:bg-purple-700 flex items-center text-sm"
                        >
                            <Upload className="mr-2" size={16} /> Tải mẫu mới (.docx)
                        </button>
                        <button 
                            onClick={() => handleDownloadGuide('invitation')}
                            className="bg-white text-purple-700 border border-purple-300 px-4 py-2 rounded shadow-sm hover:bg-purple-50 flex items-center text-sm"
                        >
                            <FileQuestion className="mr-2" size={16} /> Tải hướng dẫn biến mẫu
                        </button>
                    </div>
                </div>

                <div className="border rounded bg-gray-50 p-4">
                    <h4 className="font-bold text-gray-700 mb-2 text-sm">Danh sách mẫu Thư mời hiện có:</h4>
                    {templates.filter(t => t.type === 'invitation_word').length === 0 ? (
                        <p className="text-gray-400 italic text-sm">Chưa có mẫu nào.</p>
                    ) : (
                        <ul className="space-y-2">
                            {templates.filter(t => t.type === 'invitation_word').map(t => (
                                <li key={t.id} className="flex justify-between items-center bg-white p-2 rounded border text-sm">
                                    <span className="flex items-center text-gray-700">
                                        <FileText size={16} className="text-purple-500 mr-2"/> {t.name}
                                    </span>
                                    <button onClick={() => deleteTemplate(t.id)} className="text-red-500 hover:text-red-700 p-1">
                                        <Trash2 size={16} />
                                    </button>
                                </li>
                            ))}
                        </ul>
                    )}
                </div>
            </div>
        </div>

        {/* Payment Template Management */}
        <div className="bg-white p-6 rounded shadow border border-gray-200">
             <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
                <FileSpreadsheet className="mr-2 text-orange-600" /> Quản lý mẫu Thanh toán (Excel)
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                     <p className="text-sm text-gray-600 mb-3">
                         Tải lên các file Excel (.xlsx) mẫu để sử dụng khi xuất phiếu thanh toán giảng dạy.
                     </p>
                     <input 
                        type="file" 
                        accept=".xlsx" 
                        ref={templatePaymentInputRef} 
                        className="hidden" 
                        onChange={(e) => handleTemplateUpload(e, 'payment_excel')} 
                    />
                    <div className="flex flex-wrap gap-2 mb-4">
                        <button 
                            onClick={() => templatePaymentInputRef.current?.click()}
                            className="bg-orange-600 text-white px-4 py-2 rounded shadow hover:bg-orange-700 flex items-center text-sm"
                        >
                            <Upload className="mr-2" size={16} /> Tải mẫu mới (.xlsx)
                        </button>
                        <button 
                            onClick={() => handleDownloadGuide('payment')}
                            className="bg-white text-orange-700 border border-orange-300 px-4 py-2 rounded shadow-sm hover:bg-orange-50 flex items-center text-sm"
                        >
                            <FileQuestion className="mr-2" size={16} /> Tải hướng dẫn biến mẫu
                        </button>
                    </div>
                </div>

                <div className="border rounded bg-gray-50 p-4">
                    <h4 className="font-bold text-gray-700 mb-2 text-sm">Danh sách mẫu Excel hiện có:</h4>
                    {templates.filter(t => t.type === 'payment_excel').length === 0 ? (
                        <p className="text-gray-400 italic text-sm">Chưa có mẫu nào.</p>
                    ) : (
                        <ul className="space-y-2">
                            {templates.filter(t => t.type === 'payment_excel').map(t => (
                                <li key={t.id} className="flex justify-between items-center bg-white p-2 rounded border text-sm">
                                    <span className="flex items-center text-gray-700">
                                        <FileSpreadsheet size={16} className="text-blue-500 mr-2"/> {t.name}
                                    </span>
                                    <button onClick={() => deleteTemplate(t.id)} className="text-red-500 hover:text-red-700 p-1">
                                        <Trash2 size={16} />
                                    </button>
                                </li>
                            ))}
                        </ul>
                    )}
                </div>
            </div>
        </div>

        {/* Student List Template Management */}
        <div className="bg-white p-6 rounded shadow border border-gray-200">
             <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center">
                <FileSpreadsheet className="mr-2 text-green-600" /> Quản lý mẫu Danh sách lớp (Excel)
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                     <p className="text-sm text-gray-600 mb-3">
                         Tải lên các file Excel (.xlsx) mẫu để sử dụng khi xuất danh sách lớp sinh viên.
                     </p>
                     <input 
                        type="file" 
                        accept=".xlsx" 
                        ref={templateStudentInputRef} 
                        className="hidden" 
                        onChange={(e) => handleTemplateUpload(e, 'student_list_excel')} 
                    />
                     <div className="flex flex-wrap gap-2 mb-4">
                        <button 
                            onClick={() => templateStudentInputRef.current?.click()}
                            className="bg-green-600 text-white px-4 py-2 rounded shadow hover:bg-green-700 flex items-center text-sm"
                        >
                            <Upload className="mr-2" size={16} /> Tải mẫu mới (.xlsx)
                        </button>
                        <button 
                            onClick={() => handleDownloadGuide('student')}
                            className="bg-white text-green-700 border border-green-300 px-4 py-2 rounded shadow-sm hover:bg-green-50 flex items-center text-sm"
                        >
                            <FileQuestion className="mr-2" size={16} /> Tải hướng dẫn biến mẫu
                        </button>
                    </div>
                </div>

                <div className="border rounded bg-gray-50 p-4">
                    <h4 className="font-bold text-gray-700 mb-2 text-sm">Danh sách mẫu Excel hiện có:</h4>
                    {templates.filter(t => t.type === 'student_list_excel').length === 0 ? (
                        <p className="text-gray-400 italic text-sm">Chưa có mẫu nào.</p>
                    ) : (
                        <ul className="space-y-2">
                            {templates.filter(t => t.type === 'student_list_excel').map(t => (
                                <li key={t.id} className="flex justify-between items-center bg-white p-2 rounded border text-sm">
                                    <span className="flex items-center text-gray-700">
                                        <FileSpreadsheet size={16} className="text-green-500 mr-2"/> {t.name}
                                    </span>
                                    <button onClick={() => deleteTemplate(t.id)} className="text-red-500 hover:text-red-700 p-1">
                                        <Trash2 size={16} />
                                    </button>
                                </li>
                            ))}
                        </ul>
                    )}
                </div>
            </div>
        </div>
    </div>
  );
};

export default SystemManager;