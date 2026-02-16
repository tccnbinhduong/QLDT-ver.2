import React, { useRef } from 'react';
import { useApp } from '../store/AppContext';
import { FileText, FileSpreadsheet, File, Trash2, Download, Upload, FolderOpen } from 'lucide-react';
import { format } from 'date-fns';
import { DocumentItem } from '../types';

const DocumentManager: React.FC = () => {
  const { documents, addDocument, deleteDocument } = useApp();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const getFileIcon = (type: string) => {
    switch (type) {
      case 'word': return <FileText className="text-blue-600" size={24} />;
      case 'excel': return <FileSpreadsheet className="text-green-600" size={24} />;
      case 'pdf': return <File className="text-red-600" size={24} />;
      default: return <File className="text-gray-500" size={24} />;
    }
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return bytes + ' B';
    else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    else return (bytes / 1048576).toFixed(1) + ' MB';
  };

  const getFileType = (fileName: string): 'word' | 'excel' | 'pdf' | 'other' => {
    const ext = fileName.split('.').pop()?.toLowerCase();
    if (ext === 'doc' || ext === 'docx') return 'word';
    if (ext === 'xls' || ext === 'xlsx') return 'excel';
    if (ext === 'pdf') return 'pdf';
    return 'other';
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Limit size to 2MB because we use localStorage
    if (file.size > 2 * 1024 * 1024) {
      alert("File quá lớn! Vui lòng chọn file dưới 2MB.");
      if (fileInputRef.current) fileInputRef.current.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      const content = evt.target?.result as string;
      const newDoc: Omit<DocumentItem, 'id'> = {
        name: file.name,
        type: getFileType(file.name),
        size: file.size,
        uploadDate: new Date().toISOString(),
        content: content
      };
      addDocument(newDoc);
      if (fileInputRef.current) fileInputRef.current.value = '';
    };
    reader.readAsDataURL(file);
  };

  const handleDownload = (doc: DocumentItem) => {
    const link = document.createElement('a');
    link.href = doc.content;
    link.download = doc.name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleDelete = (id: string, name: string) => {
    if (window.confirm(`Bạn có chắc chắn muốn xóa tài liệu: ${name}?`)) {
      deleteDocument(id);
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <h1 className="text-2xl font-bold text-gray-800 flex items-center">
          <FolderOpen className="mr-3 text-yellow-600" /> Hồ sơ - Tài liệu
        </h1>
        
        <div>
          <input 
            type="file" 
            ref={fileInputRef} 
            className="hidden" 
            accept=".doc,.docx,.xls,.xlsx,.pdf,.txt"
            onChange={handleFileUpload}
          />
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 shadow"
          >
            <Upload size={18} /> Tải tài liệu lên
          </button>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow border border-gray-200 overflow-hidden">
        {documents.length === 0 ? (
          <div className="p-12 text-center flex flex-col items-center justify-center text-gray-400">
             <FolderOpen size={48} className="mb-4 text-gray-300" />
             <p className="text-lg font-medium">Chưa có tài liệu nào.</p>
             <p className="text-sm mt-1">Hãy tải lên các file Word, Excel, PDF để lưu trữ.</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left">
              <thead className="bg-gray-50 border-b">
                <tr>
                  <th className="p-4 text-gray-600 font-semibold text-sm w-16 text-center">#</th>
                  <th className="p-4 text-gray-600 font-semibold text-sm">Tên tài liệu</th>
                  <th className="p-4 text-gray-600 font-semibold text-sm">Ngày tải lên</th>
                  <th className="p-4 text-gray-600 font-semibold text-sm">Kích thước</th>
                  <th className="p-4 text-gray-600 font-semibold text-sm text-center">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {documents.map((doc) => (
                  <tr key={doc.id} className="hover:bg-gray-50 transition">
                    <td className="p-4 text-center">
                       {getFileIcon(doc.type)}
                    </td>
                    <td className="p-4">
                      <div className="font-medium text-gray-800">{doc.name}</div>
                    </td>
                    <td className="p-4 text-gray-600 text-sm">
                      {format(new Date(doc.uploadDate), 'dd/MM/yyyy HH:mm')}
                    </td>
                    <td className="p-4 text-gray-600 text-sm">
                      {formatSize(doc.size)}
                    </td>
                    <td className="p-4">
                      <div className="flex justify-center gap-3">
                        <button 
                          type="button"
                          onClick={() => handleDownload(doc)}
                          className="text-blue-600 hover:text-blue-800 p-2 rounded hover:bg-blue-50 transition-colors"
                          title="Tải xuống"
                        >
                          <Download size={18} />
                        </button>
                        <button 
                          type="button"
                          onClick={() => handleDelete(doc.id, doc.name)}
                          className="text-red-500 hover:text-red-700 p-2 rounded hover:bg-red-50 transition-colors"
                          title="Xóa"
                        >
                          <Trash2 size={18} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};

export default DocumentManager;