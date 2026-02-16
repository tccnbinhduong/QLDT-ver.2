
import React, { useState, useMemo } from 'react';
import { useApp } from '../store/AppContext';
import { Plus, Trash2, Calendar, CalendarOff, Edit2, Filter, Save, X } from 'lucide-react';
import { format } from 'date-fns';
import { parseLocal } from '../utils';
import { Holiday } from '../types';

const HolidayManager: React.FC = () => {
  const { holidays, addHoliday, updateHoliday, deleteHoliday } = useApp();
  
  // State for Form
  const [formData, setFormData] = useState({
    name: '',
    startDate: '',
    endDate: ''
  });
  const [editingId, setEditingId] = useState<string | null>(null);

  // State for Filtering
  const [selectedYear, setSelectedYear] = useState<number>(new Date().getFullYear());

  const handleSave = () => {
    if (!formData.name || !formData.startDate || !formData.endDate) {
      alert('Vui lòng điền đầy đủ thông tin: Tên đợt nghỉ, ngày bắt đầu và ngày kết thúc.');
      return;
    }
    
    if (formData.startDate > formData.endDate) {
      alert('Ngày kết thúc phải lớn hơn hoặc bằng ngày bắt đầu.');
      return;
    }

    if (editingId) {
        updateHoliday(editingId, formData);
        setEditingId(null);
    } else {
        addHoliday(formData);
    }
    
    setFormData({ name: '', startDate: '', endDate: '' });
  };

  const handleEdit = (holiday: Holiday) => {
      setFormData({
          name: holiday.name,
          startDate: holiday.startDate,
          endDate: holiday.endDate
      });
      setEditingId(holiday.id);
  };

  const handleCancelEdit = () => {
      setFormData({ name: '', startDate: '', endDate: '' });
      setEditingId(null);
  };

  // Get unique years from holidays for filter dropdown
  const uniqueYears = useMemo(() => {
      const years = new Set<number>();
      years.add(new Date().getFullYear()); // Always include current year
      holidays.forEach(h => {
          years.add(parseLocal(h.startDate).getFullYear());
          years.add(parseLocal(h.endDate).getFullYear());
      });
      return Array.from(years).sort((a, b) => b - a); // Descending
  }, [holidays]);

  // Filter and Sort holidays
  const filteredHolidays = useMemo(() => {
      return holidays
        .filter(h => {
            const startYear = parseLocal(h.startDate).getFullYear();
            const endYear = parseLocal(h.endDate).getFullYear();
            // Show if it overlaps with selected year
            return startYear === selectedYear || endYear === selectedYear;
        })
        .sort((a, b) => new Date(a.startDate).getTime() - new Date(b.startDate).getTime());
  }, [holidays, selectedYear]);

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-bold text-gray-800 flex items-center">
        <CalendarOff className="mr-3 text-red-600" /> Quản lý Ngày nghỉ (Lễ/Tết)
      </h1>
      
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {/* Form */}
        <div className={`p-6 rounded-xl shadow border h-fit transition-colors ${editingId ? 'bg-orange-50 border-orange-200' : 'bg-white border-gray-200'}`}>
          <h2 className={`text-lg font-bold mb-4 flex items-center ${editingId ? 'text-orange-700' : 'text-gray-800'}`}>
            {editingId ? <Edit2 className="mr-2" size={20} /> : <Plus className="mr-2 text-blue-600" size={20} />} 
            {editingId ? 'Cập nhật đợt nghỉ' : 'Thêm đợt nghỉ mới'}
          </h2>
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Tên đợt nghỉ (Ví dụ: Tết Nguyên Đán)</label>
              <input 
                type="text" 
                className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none" 
                placeholder="Nhập tên..."
                value={formData.name}
                onChange={e => setFormData({...formData, name: e.target.value})}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Từ ngày</label>
              <input 
                type="date" 
                className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none" 
                value={formData.startDate}
                onChange={e => setFormData({...formData, startDate: e.target.value})}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Đến ngày</label>
              <input 
                type="date" 
                className="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none" 
                value={formData.endDate}
                onChange={e => setFormData({...formData, endDate: e.target.value})}
              />
            </div>
            
            <div className="flex gap-2">
                {editingId && (
                    <button 
                        onClick={handleCancelEdit}
                        className="flex-1 bg-gray-200 text-gray-700 py-2 rounded hover:bg-gray-300 transition flex items-center justify-center font-medium"
                    >
                        <X size={18} className="mr-1" /> Hủy
                    </button>
                )}
                <button 
                    onClick={handleSave}
                    className={`flex-1 text-white py-2 rounded transition flex items-center justify-center font-medium ${editingId ? 'bg-orange-600 hover:bg-orange-700' : 'bg-blue-600 hover:bg-blue-700'}`}
                >
                    {editingId ? <><Save size={18} className="mr-2" /> Lưu</> : <><Plus size={18} className="mr-2" /> Thêm</>}
                </button>
            </div>
            
            <p className="text-xs text-gray-500 italic mt-2">
              * Lưu ý: Khi xếp lịch học, hệ thống sẽ tự động khóa các ngày nằm trong khoảng thời gian nghỉ này.
            </p>
          </div>
        </div>

        {/* List */}
        <div className="md:col-span-2 bg-white rounded-xl shadow border border-gray-200 overflow-hidden flex flex-col">
           <div className="p-4 bg-gray-50 border-b flex justify-between items-center">
             <h2 className="font-bold text-gray-700 flex items-center">
               <Calendar className="mr-2" size={20} /> Danh sách các đợt nghỉ
             </h2>
             
             {/* Year Filter */}
             <div className="flex items-center gap-2">
                 <Filter size={16} className="text-gray-500" />
                 <span className="text-sm font-medium text-gray-600">Năm:</span>
                 <select 
                    value={selectedYear}
                    onChange={(e) => setSelectedYear(Number(e.target.value))}
                    className="border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
                 >
                     {uniqueYears.map(year => (
                         <option key={year} value={year}>{year}</option>
                     ))}
                 </select>
             </div>
           </div>
           
           {filteredHolidays.length === 0 ? (
             <div className="p-8 text-center text-gray-400 italic">
               Chưa có ngày nghỉ nào trong năm {selectedYear}.
             </div>
           ) : (
             <div className="overflow-x-auto">
               <table className="w-full text-left">
                 <thead className="bg-gray-100 text-gray-600 uppercase text-xs">
                   <tr>
                     <th className="p-4 font-semibold">Tên đợt nghỉ</th>
                     <th className="p-4 font-semibold">Thời gian</th>
                     <th className="p-4 font-semibold text-center">Thao tác</th>
                   </tr>
                 </thead>
                 <tbody className="divide-y divide-gray-100">
                   {filteredHolidays.map(h => {
                     const start = format(parseLocal(h.startDate), 'dd/MM/yyyy');
                     const end = format(parseLocal(h.endDate), 'dd/MM/yyyy');
                     const isEditing = h.id === editingId;

                     return (
                       <tr key={h.id} className={`hover:bg-gray-50 transition-colors ${isEditing ? 'bg-orange-50' : ''}`}>
                         <td className="p-4 font-medium text-gray-800">
                             {h.name}
                             {isEditing && <span className="ml-2 text-xs text-orange-600 font-bold">(Đang sửa)</span>}
                         </td>
                         <td className="p-4 text-gray-600">
                           <span className="bg-blue-50 text-blue-700 px-2 py-1 rounded text-sm font-medium border border-blue-100">
                             {start} - {end}
                           </span>
                         </td>
                         <td className="p-4 text-center">
                           <div className="flex justify-center gap-2">
                               <button 
                                 onClick={() => handleEdit(h)}
                                 className="text-orange-500 hover:text-orange-700 p-2 rounded hover:bg-orange-50 transition"
                                 title="Chỉnh sửa"
                               >
                                 <Edit2 size={18} />
                               </button>
                               <button 
                                 onClick={() => {
                                   if(window.confirm(`Xóa đợt nghỉ: ${h.name}?`)) deleteHoliday(h.id);
                                 }}
                                 className="text-red-500 hover:text-red-700 p-2 rounded hover:bg-red-50 transition"
                                 title="Xóa"
                               >
                                 <Trash2 size={18} />
                               </button>
                           </div>
                         </td>
                       </tr>
                     );
                   })}
                 </tbody>
               </table>
             </div>
           )}
        </div>
      </div>
    </div>
  );
};

export default HolidayManager;
