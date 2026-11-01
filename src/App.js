import { supabase } from './supabaseClient';
import { LOCAL_ADMIN_SESSION } from './localAdmin';
import Login from './Login';
import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Download, RefreshCw, PlusCircle, Trash2, ArrowUp, ArrowDown, Eye, ArrowRightLeft, X, UserCheck, Layers, CheckSquare, Square, Search, Printer, FileText, Home, Settings, Columns, FileUp, FileDown, Type, EyeOff, Layers3, Sparkles, UserCheck2, Save, RotateCcw, BarChart2, Edit2, UserPlus, Undo, History } from 'lucide-react';

// =======================================================================
// CÁC HÀM PHỤ TRỢ & CẤU HÌNH MẶC ĐỊNH
// =======================================================================
const DEFAULT_PRINT_CONFIG = {
  headerSize: 10.5, tableSize: 10.5, footerSize: 10.5,
  headerSpacing: 2, footerSpacing: 4, fixedRowCount: 54,
  rowPadding: 0.5, marginLeft: 5, marginRight: 5, marginTop: 5, marginBottom: 5,
  headerLeftWidth: 42, footerWidth: 50, headerMinHeight: 16, footerMinHeight: 14,
  showHeader: true, showFooter: true,
  colWidths: { STT: 5, GhiChu: 18, 'HoTen': 40, 'Họ và tên': 40, 'Họ tên': 40, 'NgaySinh': 12, 'Ngày sinh': 12, 'Ngày Sinh': 12 }
};

const DEFAULT_DOC_FORMAT = {
  coQuanChuQuan: "ỦY BAN NHÂN DÂN XÃ XUÂN THỚI SƠN",
  coQuanBanHanh: "TRƯỜNG THCS TÂN XUÂN",
  diaDanh: "Xuân Thới Sơn, ngày ..... tháng ..... năm 202...",
  chucVu: "HIỆU TRƯỞNG",
  nguoiKy: "(Ký, đóng dấu và ghi rõ họ tên)"
};

const createInitialGradeState = (gradePrefix) => {
  const savedDocFormat = localStorage.getItem('app_saved_docFormat');
  const savedDocTitle = localStorage.getItem('app_saved_docTitle');
  const savedAcademicYear = localStorage.getItem('app_saved_academicYear');
  const savedPrintConfig = localStorage.getItem('app_saved_printConfig');

  return {
    gradePrefix: gradePrefix || "6",
    data: [], fileList: [], maxStudentsPerClass: 54,
    dividedClasses: null, excelColumns: [], priorityList: [],
    selectedClassIndices: [], selectedExportClasses: [], reDivideOption: 'ALL',
    printColumns: [], customColumns: [], showDuplicates: false, hiddenPrintIds: [], 
    docTitle: savedDocTitle || "DANH SÁCH HỌC SINH",
    academicYear: savedAcademicYear || "NĂM HỌC 2024 - 2025",
    docFormat: savedDocFormat ? JSON.parse(savedDocFormat) : { ...DEFAULT_DOC_FORMAT },
    printConfig: savedPrintConfig ? JSON.parse(savedPrintConfig) : { ...DEFAULT_PRINT_CONFIG },
    actionHistory: []
  };
};

const countAttr = (students, colName, value) => {
  if (!students || !Array.isArray(students)) return 0;
  return students.filter(s => (s[colName] || '').toString().trim().toUpperCase() === value.toUpperCase()).length;
};

const isYes = (val) => {
  if (!val) return false;
  const str = val.toString().trim().toUpperCase();
  return str === 'X' || str === 'CÓ' || str === '1';
};

const checkService = (st, keywords) => {
  if (!st) return false;
  for (let key in st) {
    const cleanKey = key.toString().trim().toLowerCase().normalize('NFC').replace(/\s+/g, '');
    if (keywords.some(kw => cleanKey === kw.normalize('NFC') || cleanKey.includes(kw.normalize('NFC')))) {
      return isYes(st[key]);
    }
  }
  return false;
};

const hasMayLanh = (st) => checkService(st, ['maylanh', 'máylạnh']);
const hasStem = (st) => checkService(st, ['giaoducstem', 'stem', 'giáodụcstem']);
const hasAVBN = (st) => checkService(st, ['avbn', 'anhvănbảnngữ']);
const hasKNS = (st) => checkService(st, ['kynangsong', 'kỹnăngsống', 'kns']);
const hasIC3 = (st) => checkService(st, ['ic3', 'tinhọcic3']);

const getGhiChuKey = (st) => {
  if (!st) return null;
  for (let key in st) {
    const cleanKey = key.toString().trim().toLowerCase().normalize('NFC').replace(/\s+/g, '');
    if (cleanKey === 'ghichu' || cleanKey === 'note' || cleanKey === 'ghichú') {
      return key;
    }
  }
  return null;
};

const getGhiChu = (st) => {
  if (!st) return '';
  const key = getGhiChuKey(st);
  return key ? st[key] : '';
};

const getValueScore = (val, colName) => {
  if (!val) return 0;
  const str = val.toString().trim().toUpperCase();
  if (colName === 'HL') return { 'T': 4, 'K': 3, 'Đ': 2, 'CĐ': 1 }[str] || 0;
  if (colName === 'HK') return { 'T': 3, 'K': 2, 'Đ': 1 }[str] || 0;
  if (str === 'X' || str === 'CÓ' || str === '1') return 1;
  return str;
};

const sortStudentsAlpha = (students) => {
  if (!Array.isArray(students)) return [];
  return [...students].sort((a, b) => {
    const getParts = (st) => {
      if (!st) return { ten: '', hoDem: '' };
      const name = (st.HoTen || st['Họ và tên'] || st['Họ tên'] || '').toString().trim();
      if (!name) return { ten: '', hoDem: '' };
      const p = name.split(/\s+/);
      const ten = p.pop() || ''; 
      const hoDem = p.join(' ');
      return { ten, hoDem };
    };
    const aParts = getParts(a);
    const bParts = getParts(b);
    const cmpTen = aParts.ten.localeCompare(bParts.ten, 'vi');
    if (cmpTen !== 0) return cmpTen;
    return aParts.hoDem.localeCompare(bParts.hoDem, 'vi');
  });
};

const cleanName = (name) => (name || '').toString().trim().replace(/\s+/g, ' ');

const getDob = (st) => {
  if (!st) return 'Không có DL';
  const keys = Object.keys(st);
  const dobKey = keys.find(k => {
    const lk = k.toLowerCase().normalize('NFC');
    return lk.includes('sinh') || lk.includes('ntns') || lk === 'dob' || lk === 'ngaysinh';
  });
  if (dobKey && st[dobKey] !== undefined && st[dobKey] !== null && st[dobKey] !== '') {
    return st[dobKey].toString().split('T')[0]; 
  }
  return 'Không có DL';
};

const getStudentId = (st) => {
  if (!st) return '';
  if (st._id) return st._id;
  const name = st.HoTen || st['Họ và tên'] || st['Họ tên'] || '';
  return `${name}_${getDob(st)}_${st.GioiTinh || st['Giới tính'] || ''}`;
};

const formatStatText = (text, type) => {
  if (!text) return null;
  const parts = text.split(' - ');
  return parts.map((part, idx) => {
    const m = part.match(/([A-ZĐ]+)\((\d+)\)/);
    if (m) {
       let numberColor = type === 'HL' ? '#dc2626' : '#1d4ed8'; 
       return (
         <React.Fragment key={idx}>
           {m[1]}(<span style={{ color: numberColor, fontWeight: 'bold' }}>{m[2]}</span>){idx < parts.length - 1 ? ' - ' : ''}
         </React.Fragment>
       );
    }
    return part + (idx < parts.length - 1 ? ' - ' : '');
  });
};

// =======================================================================
// COMPONENT CHÍNH CỦA ỨNG DỤNG
// =======================================================================
export default function App() {
  // 1. STATE ĐĂNG NHẬP & BẢO VỆ DỮ LIỆU
  const [authSession, setSession] = useState(null);
  const [cloudRole, setUserRole] = useState('viewer');
  const [isLocalAdmin, setIsLocalAdmin] = useState(false);
  const session = isLocalAdmin ? LOCAL_ADMIN_SESSION : authSession;
  const userRole = isLocalAdmin ? 'admin' : cloudRole;
  const [isCheckingAuth, setIsCheckingAuth] = useState(true);
  const [isDataLoaded, setIsDataLoaded] = useState(false); 

  // ---> STATE DÀNH CHO TRANG TRA CỨU PUBLIC <---
  const [publicSearchName, setPublicSearchName] = useState('');
  const [publicSearchDob, setPublicSearchDob] = useState('');
  const [publicSearchResult, setPublicSearchResult] = useState(null);
  const [showAdminLogin, setShowAdminLogin] = useState(false);

  // 2. CÁC STATE GIAO DIỆN ADMIN
  const [currentView, setCurrentView] = useState('MAIN');
  const [activeGrade, setActiveGrade] = useState("6");
  const [showServiceCol, setShowServiceCol] = useState(true);
  const [searchQuery, setSearchQuery] = useState("");
  const [newStudentData, setNewStudentData] = useState({});
  const [showStatsModal, setShowStatsModal] = useState(false);
  const [showGhiChuModal, setShowGhiChuModal] = useState(false);
  const [noteSearchQuery, setNoteSearchQuery] = useState("");
  const [noteSortOption, setNoteSortOption] = useState("NAME_ASC");
  const [unassignedModalOpen, setUnassignedModalOpen] = useState(false);
  const [studentDetailModal, setStudentDetailModal] = useState({ isOpen: false, classIdx: null, studentIdx: null, student: null });
  const [editingClassName, setEditingClassName] = useState({ classIdx: null, newName: '' });
  const [selectedMixGroups, setSelectedMixGroups] = useState([]);
  const [printHideModalOpen, setPrintHideModalOpen] = useState(false);
  const [showColMenu, setShowColMenu] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [swapModal, setSwapModal] = useState({ isOpen: false, sourceClassIdx: null, studentIdx: null, student: null, targetClassIdx: 0, targetStudentIdx: -1 });

  const jsonInputRef = useRef(null);
  const reImportExcelRef = useRef(null);
// --- BẮT ĐẦU: KHAI BÁO BIẾN CHO FILE GVCN ---
  const gvcnInputRef = useRef(null); 
  const [gvcnList, setGvcnList] = useState(() => {
    const saved = localStorage.getItem('app_gvcn_list');
    return saved ? JSON.parse(saved) : {};
  });
  // --- KẾT THÚC: KHAI BÁO BIẾN CHO FILE GVCN ---
  // 3. KHAI BÁO GRADES (Lấy ngay từ LocalStorage để không bị trắng màn hình)
  const [grades, setGrades] = useState(() => {
    const savedWorkspace = localStorage.getItem('app_workspace_data');
    if (savedWorkspace) {
      try { return JSON.parse(savedWorkspace); } catch(e){}
    }
    return {
      "6": createInitialGradeState("6"),
      "7": createInitialGradeState("7"),
      "8": createInitialGradeState("8"),
      "9": createInitialGradeState("9")
    };
  });

  // TÍNH TỔNG HỌC SINH TOÀN TRƯỜNG
  const totalStudentsWholeSchool = useMemo(() => {
    let total = 0;
    ['6', '7', '8', '9'].forEach(gKey => {
      const gData = grades[gKey];
      if (gData && gData.data) {
        total += gData.data.length;
      }
    });
    return total;
  }, [grades]);

  // STATE NHẬT KÝ THAO TÁC (LOG)
  const [activityLogs, setActivityLogs] = useState(() => {
    const savedLogs = localStorage.getItem('app_activity_logs');
    return savedLogs ? JSON.parse(savedLogs) : [];
  });
  const [showLogModal, setShowLogModal] = useState(false);

  const addLog = (actionContent) => {
    const newLog = {
      id: Date.now(),
      time: new Date().toLocaleString('vi-VN'),
      action: actionContent,
      user: session?.user?.email || 'Quản trị viên'
    };
    setActivityLogs(prev => {
      const updatedLogs = [newLog, ...prev].slice(0, 150);
      localStorage.setItem('app_activity_logs', JSON.stringify(updatedLogs));
      return updatedLogs;
    });
  };

  // 4. KIỂM TRA ĐĂNG NHẬP & KÉO DỮ LIỆU TỪ MÂY
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      fetchUserRoleAndData(session?.user?.id);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      fetchUserRoleAndData(session?.user?.id);
    });

    return () => subscription.unsubscribe();
  }, []);

  const fetchUserRoleAndData = async (userId) => {
    if (userId) {
      const { data: roleData } = await supabase.from('user_roles').select('role').eq('id', userId).single();
      if (roleData) setUserRole(roleData.role);
    } else {
      setUserRole('guest');
    }

    const { data: cloudData } = await supabase.from('app_data').select('grades_data').eq('id', 1).single();
    if (cloudData && cloudData.grades_data && Object.keys(cloudData.grades_data).length > 0) {
       setGrades(cloudData.grades_data);
       localStorage.setItem('app_workspace_data', JSON.stringify(cloudData.grades_data));
    }
    
    setIsCheckingAuth(false);
    setIsDataLoaded(true);
  };

  // 5. TỰ ĐỘNG ĐỒNG BỘ LÊN MÂY (CHỈ ADMIN)
  useEffect(() => {
    if (isCheckingAuth || !isDataLoaded) return;
    try {
      const dataToSave = JSON.parse(JSON.stringify(grades));
      Object.keys(dataToSave).forEach(key => { dataToSave[key].actionHistory = []; });
      localStorage.setItem('app_workspace_data', JSON.stringify(dataToSave));
    } catch (e) {}

    if (userRole === 'admin') {
      const timeoutId = setTimeout(async () => {
        const dataToSave = JSON.parse(JSON.stringify(grades));
        Object.keys(dataToSave).forEach(key => { dataToSave[key].actionHistory = []; });
        await supabase.from('app_data').update({ grades_data: dataToSave, updated_at: new Date() }).eq('id', 1);
      }, 3000);
      return () => clearTimeout(timeoutId);
    }
  }, [grades, isCheckingAuth, userRole, isDataLoaded]);

  // ---> HÀM XỬ LÝ KHI KHÁCH BẤM TRA CỨU <---
  const handlePublicSearch = (e) => {
    e.preventDefault();
    if (!publicSearchName.trim()) {
       setPublicSearchResult({ error: "Vui lòng nhập Họ và tên học sinh!" });
       return;
    }
    
    const queryName = publicSearchName.toLowerCase().trim();
    const queryDob = publicSearchDob.trim();
    const results = [];

    ['6', '7', '8', '9'].forEach(gKey => {
      const gData = grades[gKey];
      if (gData && gData.dividedClasses) {
        gData.dividedClasses.forEach(cls => {
          (cls.students || []).forEach(st => {
            const name = (st.HoTen || st['Họ và tên'] || '').toString().toLowerCase();
            const dob = getDob(st);
            
            if (name.includes(queryName)) {
              if (!queryDob || dob.includes(queryDob)) {
                 results.push({ 
                   name: st.HoTen || st['Họ và tên'], 
                   className: cls.className, 
                   dob: dob, 
                   gender: st.GioiTinh 
                 });
              }
            }
          });
        });
      }
    });

    if (results.length > 0) {
       setPublicSearchResult({ success: true, data: results });
    } else {
       setPublicSearchResult({ error: "❌ Không tìm thấy học sinh! Vui lòng kiểm tra lại chính tả (Ví dụ: thử gõ tên không dấu, hoặc kiểm tra lại ngày sinh)." });
    }
  };
  
  // 6. CẤU HÌNH ACTIVE
  const activeState = grades[activeGrade] || createInitialGradeState(activeGrade);
  const printConfig = activeState.printConfig || DEFAULT_PRINT_CONFIG;
  const docFormat = activeState.docFormat || DEFAULT_DOC_FORMAT;

  useEffect(() => {
    if (currentView === 'PRINT') { document.title = `Danh_Sach_Phan_Lop_Khoi_${activeGrade}`; } 
    else { document.title = "Hệ Thống Xếp Lớp Trường THCS Tân Xuân"; }
  }, [currentView, activeGrade]);

 const updateActiveGradeState = (updater) => {
    setGrades(prev => {
      const currentState = prev[activeGrade] || createInitialGradeState(activeGrade);
      
      const historyEntry = {
        timestamp: new Date().toISOString(),
        stateStr: JSON.stringify(currentState)
      };
      
      const partialState = typeof updater === 'function' ? updater(currentState) : updater;
      const updatedState = { ...currentState, ...partialState };
      
      let newHistory = currentState.actionHistory || [];
      if (partialState.data || partialState.dividedClasses) {
         newHistory = [...newHistory.slice(-19), historyEntry];
      }
      
      updatedState.actionHistory = newHistory;
      return { ...prev, [activeGrade]: updatedState };
    });
  };

  const undoLastAction = () => {
    setGrades(prev => {
      const currentState = prev[activeGrade];
      if (!currentState.actionHistory || currentState.actionHistory.length === 0) {
        alert("Không có thao tác nào để hoàn tác!");
        addLog(`[Khối ${activeGrade}] 🔙 Hoàn tác (Undo) lại thao tác trước đó`);
        return prev;
      }
      
      const newHistory = [...currentState.actionHistory];
      const lastSnapshot = newHistory.pop();
      const restoredState = JSON.parse(lastSnapshot.stateStr);
      restoredState.actionHistory = newHistory;

      return {
        ...prev,
        [activeGrade]: restoredState
      };
    });
  };

  const classesToExport = useMemo(() => {
    return (activeState.dividedClasses || []).filter((_, idx) => (activeState.selectedExportClasses || []).includes(idx));
  }, [activeState.dividedClasses, activeState.selectedExportClasses]);

  const searchResults = useMemo(() => {
    if (!searchQuery.trim() || !activeState.dividedClasses) return [];
    const query = searchQuery.toLowerCase();
    const results = [];
    activeState.dividedClasses.forEach(cls => {
      (cls.students || []).forEach((st) => {
        const name = (st.HoTen || st['Họ và tên'] || '').toString().toLowerCase();
        if (name.includes(query)) {
          results.push({ name: st.HoTen || st['Họ và tên'], className: cls.className, gender: st.GioiTinh, dob: getDob(st) });
        }
      });
    });
    return results;
  }, [searchQuery, activeState.dividedClasses]);

  const { duplicateGroups, duplicateColorMap } = useMemo(() => {
    const map = {};
    const studentData = activeState?.data || [];
    studentData.forEach(st => {
      const n = cleanName(st.HoTen || st['Họ và tên']);
      if(n) { if(!map[n]) map[n] = []; map[n].push(st); }
    });
    const dups = {}; const colors = {}; let colorIdx = 0;
    const palette = ['#e63946', '#2563eb', '#16a34a', '#d97706', '#9333ea', '#db2777', '#0891b2', '#ca8a04'];
    for (let n in map) { 
      if (map[n].length > 1) { 
        dups[n] = map[n]; colors[n] = palette[colorIdx % palette.length]; colorIdx++;
      } 
    }
    return { duplicateGroups: dups, duplicateColorMap: colors };
  }, [activeState?.data]);

  const { unassignedStudents, unassignedByLopCu } = useMemo(() => {
    const assignedIds = new Set();
    (activeState.dividedClasses || []).forEach(c => (c.students || []).forEach(s => assignedIds.add(getStudentId(s))));
    const unassigned = (activeState.data || []).filter(s => !assignedIds.has(getStudentId(s)));
    const groups = {};
    unassigned.forEach(s => {
      const l = s.LopCu || s['Lớp Cũ'] || 'NONE';
      if (!groups[l]) groups[l] = [];
      groups[l].push(s);
    });
    return { unassignedStudents: unassigned, unassignedByLopCu: groups };
  }, [activeState.data, activeState.dividedClasses]);

  const saveGlobalConfig = () => {
    try {
      localStorage.setItem('app_saved_docFormat', JSON.stringify(activeState.docFormat));
      localStorage.setItem('app_saved_docTitle', activeState.docTitle || '');
      localStorage.setItem('app_saved_academicYear', activeState.academicYear || '');
      localStorage.setItem('app_saved_printConfig', JSON.stringify(activeState.printConfig));
      alert("✅ Đã lưu cấu hình thiết lập! Các lần mở sau ứng dụng sẽ tự động áp dụng cấu hình này.");
    } catch (e) { alert("❌ Không thể lưu cấu hình vào trình duyệt!"); }
  };

  const clearGlobalConfig = () => {
    if (window.confirm("Bạn có chắc chắn muốn xóa cấu hình đã lưu và quay về mặc định?")) {
      localStorage.removeItem('app_saved_docFormat');
      localStorage.removeItem('app_saved_docTitle');
      localStorage.removeItem('app_saved_academicYear');
      localStorage.removeItem('app_saved_printConfig');
      localStorage.removeItem('app_workspace_data');
      updateActiveGradeState({
        docTitle: "DANH SÁCH HỌC SINH", academicYear: "NĂM HỌC 2024 - 2025",
        docFormat: { ...DEFAULT_DOC_FORMAT }, printConfig: { ...DEFAULT_PRINT_CONFIG }
      });
      alert("🔄 Đã khôi phục lại cấu hình mặc định ban đầu!");
    }
  };

  const handleAutoFitA4 = () => {
    const dividedClasses = activeState.dividedClasses || [];
    if (dividedClasses.length === 0) return alert(`Khối ${activeGrade} chưa có dữ liệu lớp để cân chỉnh!`);
    const maxStudentsInGrade = Math.max(...dividedClasses.map(c => (c.students || []).length), 1);
    const isHeaderHidden = activeState.printConfig?.showHeader === false;
    const maxCapacity = isHeaderHidden ? 58 : 54; 

    updateActiveGradeState(prev => ({
      printConfig: {
        ...prev.printConfig, fixedRowCount: Math.min(maxStudentsInGrade, maxCapacity),
        tableSize: 10.5, rowPadding: 0.5, headerSpacing: 2, footerSpacing: 4, marginTop: 5, marginBottom: 5
      }
    }));
    alert(`⚡ Đã cân chỉnh chuẩn cho Khối ${activeGrade}!\n(Năng suất chứa tối đa: ${maxCapacity} HS/trang A4)`);
  };

  const checkDuplicateInClasses = () => {
    const dividedClasses = activeState.dividedClasses || [];
    if (!dividedClasses || dividedClasses.length === 0) return alert(`Khối ${activeGrade} chưa được phân lớp!`);

    const duplicatesFound = [];
    dividedClasses.forEach((cls) => {
      const nameCounts = {};
      (cls.students || []).forEach(st => {
        const name = cleanName(st.HoTen || st['Họ và tên']);
        if (name) {
          if (!nameCounts[name]) nameCounts[name] = [];
          nameCounts[name].push(st);
        }
      });
      for (let name in nameCounts) {
        if (nameCounts[name].length > 1) {
          duplicatesFound.push({
            className: cls.className,
            studentName: nameCounts[name][0].HoTen || nameCounts[name][0]['Họ và tên'],
            count: nameCounts[name].length
          });
        }
      }
    });

    if (duplicatesFound.length === 0) {
      alert(`✅ XÁC NHẬN: Khối ${activeGrade} KHÔNG CÓ học sinh nào bị trùng Họ và Tên trong cùng một lớp!`);
    } else {
      let msg = `⚠️ PHÁT HIỆN HỌC SINH TRÙNG HỌ TÊN TRONG CÙNG LỚP (Khối ${activeGrade}):\n\n`;
      duplicatesFound.forEach((item, idx) => {
        msg += `${idx + 1}. Học sinh "${item.studentName}" bị trùng (${item.count} em) tại: ${item.className}\n`;
      });
      msg += `\n👉 Báo ý: Thầy nên dùng nút "✏️" của học sinh đó để chuyển 1 bạn sang lớp khác!`;
      alert(msg);
    }
  };

  const handleToggleExportClass = (idx) => {
    updateActiveGradeState(prev => {
      const current = prev.selectedExportClasses || [];
      if (current.includes(idx)) return { selectedExportClasses: current.filter(i => i !== idx) };
      return { selectedExportClasses: [...current, idx].sort((a,b)=>a-b) };
    });
  };

  const handleToggleAllExportClasses = (selectAll) => {
    updateActiveGradeState(prev => ({ selectedExportClasses: selectAll ? (prev.dividedClasses || []).map((_, i) => i) : [] }));
  };

  const exportOriginalDataToExcel = () => {
    if (!activeState.data || activeState.data.length === 0) {
      return alert("Chưa có dữ liệu gốc để xuất!");
    }

    const studentClassMap = {};
    if (activeState.dividedClasses) {
      activeState.dividedClasses.forEach(cls => {
        (cls.students || []).forEach(st => {
          const stId = getStudentId(st);
          studentClassMap[stId] = cls.className;
        });
      });
    }

    let globalStt = 1;
    const exportData = activeState.data.map(st => {
      let cleanStudent = { ...st };
      delete cleanStudent._id;
      delete cleanStudent.STT;
      delete cleanStudent['SỐ THỨ TỰ'];
      delete cleanStudent['Số Thứ Tự'];
      
      const stId = getStudentId(st);
      const assignedClass = studentClassMap[stId] || "Chưa xếp lớp";

      return {
        "STT": globalStt++,
        "Lớp mới": "'" + assignedClass,
        ...cleanStudent
      };
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);
    XLSX.utils.book_append_sheet(wb, ws, `DuLieuGoc_Khoi_${activeGrade}`);
    XLSX.writeFile(wb, `KiemTra_DuLieuGoc_Khoi_${activeGrade}.xlsx`);
  };
// --- BẮT ĐẦU: HÀM XUẤT EXCEL THỐNG KÊ (CÓ HSHN + GVCN) ---
  const exportStatsToExcel = () => {
    if (classesToExport.length === 0) return alert("Vui lòng chọn ít nhất 1 lớp để xuất báo cáo!");

    const exportData = [];
    let totalSiSo = 0;
    let totalNam = 0;
    let totalNu = 0;
    let totalMayLanh = 0;
    let totalKNS = 0;
    let totalStem = 0;
    let totalAVBN = 0;
    let totalIC3 = 0;
    let totalHSHN = 0;

    classesToExport.forEach((cls, idx) => {
      const visibleStudents = (cls.students || []).filter(s => !s.__isEmptyRow && !(activeState.hiddenPrintIds || []).includes(getStudentId(s)));
      const nam = visibleStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() === 'nam').length;
      const nu = visibleStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() !== 'nam').length;
      
      const mayLanh = visibleStudents.filter(hasMayLanh).length;
      const kns = visibleStudents.filter(hasKNS).length;
      const stem = visibleStudents.filter(hasStem).length;
      const avbn = visibleStudents.filter(hasAVBN).length;
      const ic3 = visibleStudents.filter(hasIC3).length;
      
      // Đếm số lượng học sinh có ghi chú chứa chữ "HSHN"
      const hshn = visibleStudents.filter(st => {
         const note = getGhiChu(st);
         return note && note.toString().toUpperCase().includes('HSHN');
      }).length;

      totalSiSo += visibleStudents.length;
      totalNam += nam;
      totalNu += nu;
      totalMayLanh += mayLanh;
      totalKNS += kns;
      totalStem += stem;
      totalAVBN += avbn;
      totalIC3 += ic3;
      totalHSHN += hshn;
      
      // Lấy tên GVCN từ dữ liệu đã nạp (nếu lớp chưa có GVCN thì để trống)
      const gvcn = gvcnList[cls.className] || "";

      exportData.push({
        "STT": idx + 1,
        "TÊN LỚP": cls.className,
        "GVCN": gvcn,
        "SỈ SỐ": visibleStudents.length,
        "NAM": nam,
        "NỮ": nu,
        "MÁY LẠNH": mayLanh,
        "KỸ NĂNG SỐNG": kns,
        "GIÁO DỤC STEM": stem,
        "AVBN": avbn,
        "IC3": ic3,
        "HSHN": hshn,
        "GHI CHÚ": ""
      });
    });

    exportData.push({
      "STT": "",
      "TÊN LỚP": "TỔNG CỘNG",
      "GVCN": "",
      "SỈ SỐ": totalSiSo,
      "NAM": totalNam,
      "NỮ": totalNu,
      "MÁY LẠNH": totalMayLanh,
      "KỸ NĂNG SỐNG": totalKNS,
      "GIÁO DỤC STEM": totalStem,
      "AVBN": totalAVBN,
      "IC3": totalIC3,
      "HSHN": totalHSHN,
      "GHI CHÚ": ""
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);

    const colWidths = [
      { wch: 5 },  // STT
      { wch: 10 }, // TÊN LỚP
      { wch: 25 }, // GVCN
      { wch: 10 }, // SỈ SỐ
      { wch: 8 },  // NAM
      { wch: 8 },  // NỮ
      { wch: 13 }, // MÁY LẠNH
      { wch: 16 }, // KỸ NĂNG SỐNG
      { wch: 16 }, // GIÁO DỤC STEM
      { wch: 12 }, // AVBN
      { wch: 10 }, // IC3
      { wch: 10 }, // HSHN
      { wch: 20 }  // GHI CHÚ
    ];
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, `ThongKe_SiSo_Khoi_${activeGrade}`);
    XLSX.writeFile(wb, `Thong_Ke_Si_So_Khoi_${activeGrade}.xlsx`);
  };
  // --- KẾT THÚC: HÀM XUẤT EXCEL THỐNG KÊ ---

  const exportToExcel = () => {
    if (classesToExport.length === 0) return alert("Vui lòng chọn ít nhất 1 lớp trong 'Chọn Lớp Xuất/In' để xuất file!");
    let allStudentsExport = [];
    let globalStt = 1;
    classesToExport.forEach((cls) => {
      const visibleStudents = (cls.students || []).filter(s => !s.__isEmptyRow && !(activeState.hiddenPrintIds || []).includes(getStudentId(s)));
      visibleStudents.forEach(student => { 
        let cleanStudent = { ...student };
        delete cleanStudent._id;
        delete cleanStudent.STT;
        delete cleanStudent['SỐ THỨ TỰ'];
        delete cleanStudent['Số Thứ Tự'];

        const noteKey = getGhiChuKey(cleanStudent);
        let noteValue = "";
        if (noteKey) {
          noteValue = cleanStudent[noteKey];
          delete cleanStudent[noteKey];
        }

        const row = { 
          "STT": globalStt++, 
          "Lớp mới": "'" + cls.className, 
          ...cleanStudent 
        };

        (activeState.customColumns || []).forEach(c => { 
          row[c.title] = student[c.id] || ""; 
        });

        if (noteKey || noteValue) {
          row["Ghi chú"] = noteValue;
        }

        allStudentsExport.push(row); 
      });
    });
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(allStudentsExport);
    XLSX.utils.book_append_sheet(wb, ws, `Khoi_${activeGrade}`);
    XLSX.writeFile(wb, `Danh_Sach_Phan_Lop_Khoi_${activeGrade}.xlsx`);
  };

  // --- BẮT ĐẦU: HÀM NẠP FILE EXCEL GVCN ---
  const handleImportGVCN = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
        
        const newGvcnMap = { ...gvcnList };
        let count = 0;
        
        data.forEach(row => {
          const keys = Object.keys(row);
          // Tự động tìm cột chứa chữ "LỚP" và "GVCN"
          const lopKey = keys.find(k => k.toString().toLowerCase().includes('lớp') || k.toString().toLowerCase().includes('lop'));
          const gvcnKey = keys.find(k => k.toString().toLowerCase().includes('gvcn') || k.toString().toLowerCase().includes('giáo viên'));
          
          if (lopKey && gvcnKey && row[lopKey] && row[gvcnKey]) {
            const className = row[lopKey].toString().trim();
            const teacherName = row[gvcnKey].toString().trim();
            newGvcnMap[className] = teacherName;
            count++;
          }
        });
        
        setGvcnList(newGvcnMap);
        localStorage.setItem('app_gvcn_list', JSON.stringify(newGvcnMap));
        alert(`✅ Đã nạp và lưu thành công danh sách ${count} Giáo viên chủ nhiệm!`);
      } catch (err) {
        console.error(err);
        alert("❌ Đã xảy ra lỗi khi đọc File GVCN! Vui lòng kiểm tra lại file Excel.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = null;
  };
  // --- KẾT THÚC: HÀM NẠP FILE EXCEL GVCN ---
  const exportSessionFile = () => {
    const workspaceState = { grades, activeGrade };
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(workspaceState));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", `PhienLamViec_ToanTruong_Khoi6789_${new Date().toISOString().slice(0,10)}.json`);
    document.body.appendChild(downloadAnchor); downloadAnchor.click(); downloadAnchor.remove();
  };

  const importSessionFile = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target.result);
        if (parsed.grades) {
          const mergedGrades = {};
          ['6', '7', '8', '9'].forEach(gKey => { mergedGrades[gKey] = { ...createInitialGradeState(gKey), ...(parsed.grades[gKey] || {}) }; });
          setGrades(mergedGrades);
          if (parsed.activeGrade) setActiveGrade(parsed.activeGrade);
          alert("📂 Đã mở thành công phiên làm việc 4 khối từ File!");
        } else { alert("❌ File không khớp cấu hình 4 Khối!"); }
      } catch (err) { alert("❌ File cấu hình không hợp lệ hoặc bị lỗi!"); }
    };
    reader.readAsText(file); e.target.value = null;
  };

  const handleImportExportedExcel = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result; 
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]]; 
        const parsedData = XLSX.utils.sheet_to_json(ws, { defval: "" });
        
        if (parsedData.length > 0) {
          const keys = Object.keys(parsedData[0]);
          const lopMoiKey = keys.find(k => {
             const cleanK = k.toString().trim().toLowerCase().normalize('NFC');
             return cleanK === 'lớp mới' || cleanK === 'lop moi' || cleanK === "'lớp mới" || cleanK === 'lop_moi';
          });

          if (lopMoiKey) {
            const classMap = {};
            const dataWithIds = [];

            parsedData.forEach((row, i) => {
              const stId = row._id || `st_${Date.now()}_${i}_${Math.random().toString(36).substr(2, 6)}`;
              let className = (row[lopMoiKey] || "").toString().trim();
              if (className.startsWith("'")) className = className.substring(1).trim();
              
              const student = { ...row, _id: stId };
              delete student[lopMoiKey]; 
              delete student.STT;
              delete student['SỐ THỨ TỰ'];
              delete student['Số Thứ Tự'];

              dataWithIds.push(student);

              const isUnassigned = !className || 
                                   className.toLowerCase().includes('chưa xếp') || 
                                   className.toLowerCase().includes('chưa phân') || 
                                   className.toLowerCase() === 'none' ||
                                   className.toLowerCase() === 'chưa rõ';

              if (!isUnassigned) {
                if (!classMap[className]) classMap[className] = [];
                classMap[className].push(student);
              }
            });

            const reconstructedClasses = Object.keys(classMap).map(clsName => ({ 
              className: clsName, 
              students: sortStudentsAlpha(classMap[clsName]) 
            }));

            const sampleRow = { ...dataWithIds[0] }; 
            delete sampleRow._id;
            const cols = Object.keys(sampleRow);

            const findCol = (kws) => cols.find(c => kws.includes(c.toString().trim().toLowerCase().normalize('NFC').replace(/\s+/g, '')));
            const defaultPriorities = [
               findCol(['hl']), findCol(['hk']), 
               findCol(['giaoducstem', 'stem', 'giáodụcstem']),
               findCol(['maylanh', 'máylạnh']),
               findCol(['kynangsong', 'kỹnăngsống', 'kns']),
               findCol(['avbn', 'anhvănbảnngữ']),
               findCol(['ic3', 'tinhọcic3'])
            ].filter(Boolean);

            const defaultInCols = ['HoTen', 'Họ và tên', 'NgaySinh', 'Ngày sinh', 'GioiTinh', 'Giới tính', 'HL', 'HK'].filter(c => cols.includes(c));
            const newPrintCols = defaultInCols.length > 0 ? defaultInCols : cols.slice(0, 4);

            updateActiveGradeState({
              excelColumns: cols,
              priorityList: defaultPriorities,
              printColumns: newPrintCols,
              data: dataWithIds,
              dividedClasses: reconstructedClasses,
              selectedExportClasses: reconstructedClasses.map((_, i) => i),
              fileList: [file.name],
              selectedClassIndices: [],
              showDuplicates: false,
              hiddenPrintIds: []
            });

            alert(`✅ Import ngược thành công ${parsedData.length} học sinh (${reconstructedClasses.length} lớp) cho Khối ${activeGrade}!`);
          } else { 
            alert(`❌ File Excel không chứa cột 'Lớp mới'.\nCác cột phần mềm đang đọc được từ file là: [ ${keys.join(", ")} ]\n\nVui lòng chọn đúng file Excel đã xuất từ hệ thống!`); 
          }
        } else { 
          alert("❌ File Excel trống hoặc không có dữ liệu!"); 
        }
      } catch (err) { 
        console.error(err);
        alert("❌ Đã xảy ra lỗi khi đọc File Excel!"); 
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = null;
  };

  const handleColWidthChange = (colId, value) => {
    updateActiveGradeState(prev => ({
      printConfig: { ...prev.printConfig, colWidths: { ...(prev.printConfig?.colWidths || {}), [colId]: value ? Number(value) : '' } }
    }));
  };

  const addCustomColumn = () => {
    const newId = `custom_${Date.now()}`; const newTitle = `Cột mới ${(activeState.customColumns || []).length + 1}`;
    updateActiveGradeState(prev => ({ customColumns: [...(prev.customColumns || []), { id: newId, title: newTitle }], printColumns: [...(prev.printColumns || []), newId] }));
  };

  const updateCustomColumnTitle = (id, newTitle) => {
    updateActiveGradeState(prev => ({ customColumns: (prev.customColumns || []).map(col => col.id === id ? { ...col, title: newTitle } : col) }));
  };

  const removeCustomColumn = (id) => {
    updateActiveGradeState(prev => ({ customColumns: (prev.customColumns || []).filter(col => col.id !== id), printColumns: (prev.printColumns || []).filter(c => c !== id) }));
  };

  const handleFileUpload = (e) => {
    const files = Array.from(e.target.files); if (files.length === 0) return;
    files.forEach(file => {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const bstr = evt.target.result; const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]]; const parsedData = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (parsedData.length > 0) {
          addLog(`[Khối ${activeGrade}] Nạp file Excel "${file.name}" (${parsedData.length} học sinh)`);
          const cols = Object.keys(parsedData[0]);
          updateActiveGradeState(prev => {
            const currentPriority = prev.priorityList || [];
            const findCol = (kws) => cols.find(c => kws.includes(c.toString().trim().toLowerCase().normalize('NFC').replace(/\s+/g, '')));
            const defaultPriorities = [
               findCol(['hl']), findCol(['hk']), 
               findCol(['giaoducstem', 'stem', 'giáodụcstem']),
               findCol(['maylanh', 'máylạnh']),
               findCol(['kynangsong', 'kỹnăngsống', 'kns']),
               findCol(['avbn', 'anhvănbảnngữ']),
               findCol(['ic3', 'tinhọcic3'])
            ].filter(Boolean);
            const newPriorityList = currentPriority.length === 0 ? defaultPriorities : currentPriority;
            let newPrintCols = prev.printColumns || [];
            if (newPrintCols.length === 0) {
              const defaultInCols = ['HoTen', 'Họ và tên', 'NgaySinh', 'Ngày sinh', 'GioiTinh', 'Giới tính', 'HL', 'HK'].filter(c => cols.includes(c));
              newPrintCols = defaultInCols.length > 0 ? defaultInCols : cols.slice(0, 4);
            }
            const dataWithIds = parsedData.map((r, i) => ({ ...r, _id: `st_${Date.now()}_${i}_${Math.random().toString(36).substr(2, 6)}` }));
            return {
              excelColumns: cols, priorityList: newPriorityList, printColumns: newPrintCols,
              data: [...(prev.data || []), ...dataWithIds], fileList: [...(prev.fileList || []), file.name],
              dividedClasses: prev.dividedClasses || [], selectedClassIndices: [], selectedExportClasses: [], 
              showDuplicates: false, hiddenPrintIds: prev.hiddenPrintIds || []
            };
          });
        }
      };
      reader.readAsBinaryString(file);
    });
    e.target.value = null;
  };

  const addPriority = () => {
    const excelCols = activeState.excelColumns || []; const priorityList = activeState.priorityList || [];
    const unusedCol = excelCols.find(col => !priorityList.includes(col));
    if (unusedCol) updateActiveGradeState(prev => ({ priorityList: [...(prev.priorityList || []), unusedCol] }));
  };

  const removePriority = (index) => { updateActiveGradeState(prev => { const newList = [...(prev.priorityList || [])]; newList.splice(index, 1); return { priorityList: newList }; }); };
  const updatePriority = (index, newColValue) => { updateActiveGradeState(prev => { const newList = [...(prev.priorityList || [])]; newList[index] = newColValue; return { priorityList: newList }; }); };
  const movePriority = (index, direction) => { updateActiveGradeState(prev => { const newList = [...(prev.priorityList || [])]; const targetIndex = index + direction; if (targetIndex >= 0 && targetIndex < newList.length) { const temp = newList[index]; newList[index] = newList[targetIndex]; newList[targetIndex] = temp; } return { priorityList: newList }; }); };
  const handleDeleteOriginalStudent = (indexToRemove) => {
    if (window.confirm("Thầy có chắc chắn muốn xóa học sinh này khỏi danh sách gốc không?\n(Lưu ý: Bạn nên bấm Xếp Lớp lại sau khi xóa để cập nhật kết quả)")) {
      updateActiveGradeState(prev => { 
        const newData = [...(prev.data || [])]; 
        const removedName = newData[indexToRemove]?.HoTen || newData[indexToRemove]?.['Họ và tên'] || 'Không tên';
        
        addLog(`[Khối ${activeGrade}] Xóa học sinh "${removedName}" khỏi danh sách Excel gốc`);
        
        newData.splice(indexToRemove, 1); 
        return { data: newData }; 
      });
    }
  };

  const handleAddStudentToOriginal = () => {
    if (!newStudentData['HoTen'] && !newStudentData['Họ và tên']) return alert("⚠️ Vui lòng nhập ít nhất Họ và tên của học sinh mới!");
    
    const studentName = newStudentData['HoTen'] || newStudentData['Họ và tên'];
    addLog(`[Khối ${activeGrade}] Thêm học sinh mới: "${studentName}" vào danh sách gốc`);
    
    updateActiveGradeState(prev => { 
      const newData = [...(prev.data || [])]; 
      newData.push({ ...newStudentData, _id: `st_${Date.now()}_${Math.random().toString(36).substr(2, 6)}` }); 
      return { data: newData }; 
    });
    setNewStudentData({}); alert("✅ Đã chèn thêm học sinh thành công vào danh sách gốc!");
  };

const saveClassName = () => {
    if (!editingClassName.newName.trim()) return setEditingClassName({ classIdx: null, newName: '' });
    
    updateActiveGradeState(prev => { 
      const newClasses = [...(prev.dividedClasses || [])]; 
      const oldName = newClasses[editingClassName.classIdx].className;
      addLog(`[Khối ${activeGrade}] Đổi tên lớp từ "${oldName}" thành "${editingClassName.newName}"`);
      
      newClasses[editingClassName.classIdx].className = editingClassName.newName; 
      return { dividedClasses: newClasses }; 
    });
    setEditingClassName({ classIdx: null, newName: '' });
  };

  const handleCellEdit = (rowIndex, col, value) => {
    updateActiveGradeState(prev => {
      const newData = [...(prev.data || [])];
      const oldStudent = newData[rowIndex];
      const oldStudentId = getStudentId(oldStudent);

      addLog(`[Khối ${activeGrade}] Chỉnh sửa HS "${oldStudent.HoTen || oldStudent['Họ và tên']}" (Đổi cột [${col}] thành "${value}")`);

      const updatedStudent = { ...oldStudent, [col]: value };
      newData[rowIndex] = updatedStudent;
      
      let newDividedClasses = prev.dividedClasses ? JSON.parse(JSON.stringify(prev.dividedClasses)) : null;
      if (newDividedClasses) {
        newDividedClasses.forEach(cls => {
          const sIdx = cls.students.findIndex(s => getStudentId(s) === oldStudentId);
          if (sIdx !== -1) { 
            cls.students[sIdx] = { ...cls.students[sIdx], [col]: value }; 
            const colLower = col.toLowerCase();
            if (colLower.includes('tên') || colLower.includes('hoten')) {
              cls.students = sortStudentsAlpha(cls.students);
            }
          }
        });
      }
      return { data: newData, dividedClasses: newDividedClasses };
    });
  };

  const countServicesFunc = (student) => {
    if (!student) return 0;
    const serviceCols = (activeState.priorityList || []).slice(2);
    if (serviceCols.length > 0) {
      return serviceCols.reduce((acc, col) => acc + (isYes(student[col]) ? 1 : 0), 0);
    }
    let cnt = 0;
    if (hasMayLanh(student)) cnt++;
    if (hasStem(student)) cnt++;
    if (hasAVBN(student)) cnt++;
    if (hasKNS(student)) cnt++;
    if (hasIC3(student)) cnt++;
    return cnt;
  };

  const executeDivisionOnStudents = (studentsToDivide, customStartClassNum = null) => {
    const numClasses = Math.ceil(studentsToDivide.length / (activeState.maxStudentsPerClass || 54)) || 1;
    const startNum = customStartClassNum || ((activeState.dividedClasses || []).length + 1);
    const resultClasses = Array.from({ length: numClasses }, (_, i) => ({ className: `${activeGrade}/${startNum + i}`, students: [] }));
    const priorityList = activeState.priorityList || [];
    const p1 = priorityList[0]; const p2 = priorityList[1]; 
    const groupsByService = {};
    
    studentsToDivide.forEach(s => {
      const sc = countServicesFunc(s);
      if (!groupsByService[sc]) groupsByService[sc] = [];
      groupsByService[sc].push(s);
    });

    const sortedServiceCounts = Object.keys(groupsByService).map(Number).sort((a, b) => b - a);
    sortedServiceCounts.forEach(sc => {
      const groupStudents = groupsByService[sc];
      const sortFn = (a, b) => {
        const vA1 = p1 ? getValueScore(a[p1], p1) : 0; const vB1 = p1 ? getValueScore(b[p1], p1) : 0;
        if (vB1 !== vA1) return (typeof vB1 === 'number' && typeof vA1 === 'number') ? vB1 - vA1 : String(vB1).localeCompare(String(vA1));
        const vA2 = p2 ? getValueScore(a[p2], p2) : 0; const vB2 = p2 ? getValueScore(b[p2], p2) : 0;
        if (vB2 !== vA2) return (typeof vB2 === 'number' && typeof vA2 === 'number') ? vB2 - vA2 : String(vB2).localeCompare(String(vB2));
        return 0;
      };
      const males = groupStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() === 'nam').sort(sortFn);
      const females = groupStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() !== 'nam').sort(sortFn);
      const interleaved = []; let mIdx = 0, fIdx = 0;
      while (mIdx < males.length || fIdx < females.length) {
        if (mIdx < males.length) interleaved.push(males[mIdx++]);
        if (fIdx < females.length) interleaved.push(females[fIdx++]);
      }
      let sRemaining = interleaved.length;
      const maxCap = activeState.maxStudentsPerClass || 54;
      const Q = new Array(numClasses).fill(0);
      let tempRemaining = sRemaining;
      for (let i = 0; i < numClasses; i++) {
        const remCap = maxCap - resultClasses[i].students.length;
        if (remCap > 0 && tempRemaining > 0) {
          const take = Math.min(remCap, tempRemaining);
          Q[i] = take; tempRemaining -= take;
        }
      }
      const allocated = new Array(numClasses).fill(0);
      interleaved.forEach(student => {
        let bestClass = -1; let minRatio = Infinity;
        for (let i = 0; i < numClasses; i++) {
          if (Q[i] > 0 && allocated[i] < Q[i]) {
            const ratio = allocated[i] / Q[i];
            if (ratio < minRatio) { minRatio = ratio; bestClass = i; }
          }
        }
        if (bestClass !== -1) {
          resultClasses[bestClass].students.push(student); allocated[bestClass]++;
        } else {
           let placed = false;
           for (let i = 0; i < numClasses; i++) {
             if (resultClasses[i].students.length < maxCap) { resultClasses[i].students.push(student); placed = true; break; }
           }
           if (!placed) resultClasses[numClasses - 1].students.push(student);
        }
      });
    });
    resultClasses.forEach(cls => { cls.students = sortStudentsAlpha(cls.students); });
    return resultClasses;
  };

  const handleMixSelectedGroups = () => {
    if (selectedMixGroups.length === 0) return alert("Vui lòng tích chọn ít nhất 1 nhóm/lớp cũ để Gom dịch vụ!");
    const studentsToMix = unassignedStudents.filter(s => selectedMixGroups.includes(s.LopCu || s['Lớp Cũ'] || 'NONE'));
    const newClasses = executeDivisionOnStudents(studentsToMix);
    updateActiveGradeState(prev => {
      const merged = [...(prev.dividedClasses || []), ...newClasses];
      return { dividedClasses: merged, selectedExportClasses: merged.map((_, i) => i) };
    });
    setSelectedMixGroups([]);
  };

  const handleKeepIntactGroup = (groupName) => {
    const studentsToKeep = unassignedStudents.filter(s => (s.LopCu || s['Lớp Cũ'] || 'NONE') === groupName);
    const newClassName = groupName === 'NONE' ? `${activeGrade}/Mới` : groupName.toString().replace("Lớp ", ""); 
    const newClass = { className: newClassName, students: sortStudentsAlpha(studentsToKeep) };
    updateActiveGradeState(prev => {
      const merged = [...(prev.dividedClasses || []), newClass];
      return { dividedClasses: merged, selectedExportClasses: merged.map((_, i) => i) };
    });
  };

  const processReDivisionSelected = () => {
    const dividedClasses = activeState.dividedClasses || [];
    const selectedIndices = activeState.selectedClassIndices || [];
    if (!dividedClasses || selectedIndices.length < 1) return;

    const targetIndices = [...selectedIndices].sort((a, b) => a - b);
    let studentsToReDivide = [];
    targetIndices.forEach(idx => { if (dividedClasses[idx]) studentsToReDivide = [...studentsToReDivide, ...(dividedClasses[idx].students || [])]; });

    const newClassStudents = targetIndices.map(() => []);
    const priorityList = activeState.priorityList || [];
    const p1 = priorityList[0]; const p2 = priorityList[1]; 
    const getPrimaryScore = (s) => {
      if (activeState.reDivideOption === 'ALL') return countServicesFunc(s);
      return isYes(s[activeState.reDivideOption]) ? 1 : 0;
    };

    const groups = {};
    studentsToReDivide.forEach(s => {
      const sc = getPrimaryScore(s);
      if (!groups[sc]) groups[sc] = [];
      groups[sc].push(s);
    });

    const sortedScores = Object.keys(groups).map(Number).sort((a, b) => b - a);
    sortedScores.forEach(sc => {
      const groupStudents = groups[sc];
      const sortFn = (a, b) => {
        const vA1 = p1 ? getValueScore(a[p1], p1) : 0; const vB1 = p1 ? getValueScore(b[p1], p1) : 0;
        if (vB1 !== vA1) return (typeof vB1 === 'number' && typeof vA1 === 'number') ? vB1 - vA1 : String(vB1).localeCompare(String(vA1));
        const vA2 = p2 ? getValueScore(a[p2], p2) : 0; const vB2 = p2 ? getValueScore(b[p2], p2) : 0;
        if (vB2 !== vA2) return (typeof vB2 === 'number' && typeof vA2 === 'number') ? vB2 - vA2 : String(vB2).localeCompare(String(vB2));
        return 0;
      };

      const males = groupStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() === 'nam').sort(sortFn);
      const females = groupStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() !== 'nam').sort(sortFn);
      const interleaved = []; let mIdx = 0, fIdx = 0;
      while (mIdx < males.length || fIdx < females.length) {
        if (mIdx < males.length) interleaved.push(males[mIdx++]);
        if (fIdx < females.length) interleaved.push(females[fIdx++]);
      }

      let sRemaining = interleaved.length;
      const numTargets = targetIndices.length;
      const maxCap = activeState.maxStudentsPerClass || 54;
      const Q = new Array(numTargets).fill(0);
      let tempRemaining = sRemaining;
      
      for (let i = 0; i < numTargets; i++) {
        const remCap = maxCap - newClassStudents[i].length;
        if (remCap > 0 && tempRemaining > 0) {
          const take = Math.min(remCap, tempRemaining);
          Q[i] = take; tempRemaining -= take;
        }
      }

      const allocated = new Array(numTargets).fill(0);
      interleaved.forEach(student => {
        let bestClass = -1; let minRatio = Infinity;
        for (let i = 0; i < numTargets; i++) {
          if (Q[i] > 0 && allocated[i] < Q[i]) {
            const ratio = allocated[i] / Q[i];
            if (ratio < minRatio) { minRatio = ratio; bestClass = i; }
          }
        }
        if (bestClass !== -1) {
          newClassStudents[bestClass].push(student); allocated[bestClass]++;
        } else {
           let placed = false;
           for (let i = 0; i < numTargets; i++) {
             if (newClassStudents[i].length < maxCap) { newClassStudents[i].push(student); placed = true; break; }
           }
           if (!placed) newClassStudents[numTargets - 1].push(student);
        }
      });
    });

    updateActiveGradeState(prev => {
      const updatedClasses = [...(prev.dividedClasses || [])];
      targetIndices.forEach((targetIdx, i) => { if (updatedClasses[targetIdx]) { updatedClasses[targetIdx].students = sortStudentsAlpha(newClassStudents[i]); } });
      return { dividedClasses: updatedClasses, selectedClassIndices: [] };
    });
  };

  const toggleSelectClass = (index) => {
    updateActiveGradeState(prev => {
      const selectedClassIndices = prev.selectedClassIndices || [];
      const isSelected = selectedClassIndices.includes(index);
      return { selectedClassIndices: isSelected ? selectedClassIndices.filter(i => i !== index) : [...selectedClassIndices, index] };
    });
  };

  const deleteSelectedClasses = () => {
    const selectedIndices = activeState.selectedClassIndices || [];
    if (selectedIndices.length === 0) return;
    if (window.confirm(`CẢNH BÁO Khối ${activeGrade}: Bạn đang xóa lớp khỏi danh sách hiển thị! Toàn bộ học sinh trong lớp sẽ trở về trạng thái "Chưa xếp lớp". Bạn có chắc chắn?`)) {
      
      addLog(`[Khối ${activeGrade}] Xóa ${selectedIndices.length} lớp đang chọn (Đưa toàn bộ học sinh về danh sách Chưa xếp lớp)`);
      
      updateActiveGradeState(prev => {
        const updatedClasses = (prev.dividedClasses || []).filter((_, idx) => !selectedIndices.includes(idx));
        return { dividedClasses: updatedClasses, selectedClassIndices: [], selectedExportClasses: updatedClasses.map((_, i) => i) };
      });
    }
  };

  const openStudentDetailModal = (classIdx, studentIdx, student) => { setStudentDetailModal({ isOpen: true, classIdx, studentIdx, student }); };
  const openSwapModal = (sourceClassIdx, studentIdx, student) => {
    const initialTargetIdx = sourceClassIdx === 0 ? (activeState.dividedClasses.length > 1 ? 1 : 0) : 0;
    setSwapModal({ isOpen: true, sourceClassIdx, studentIdx, student, targetClassIdx: initialTargetIdx, targetStudentIdx: -1 });
  };

  const executeManualSwap = () => {
    const dividedClasses = activeState.dividedClasses || [];
    if (!dividedClasses || dividedClasses.length === 0) return;
    const { sourceClassIdx, studentIdx, targetClassIdx, targetStudentIdx } = swapModal;
    const updated = JSON.parse(JSON.stringify(dividedClasses));
    const sourceStudent = updated[sourceClassIdx]?.students?.[studentIdx];
    if (!sourceStudent) return;

    if (targetStudentIdx === -1) {
      updated[sourceClassIdx].students.splice(studentIdx, 1);
      updated[targetClassIdx].students.push(sourceStudent);
      addLog(`[Khối ${activeGrade}] Chuyển học sinh ${sourceStudent.HoTen || sourceStudent['Họ và tên']} từ lớp ${updated[sourceClassIdx].className} sang ${updated[targetClassIdx].className}`);
    } else {
      const targetStudent = updated[targetClassIdx].students[targetStudentIdx];
      updated[sourceClassIdx].students[studentIdx] = targetStudent;
      updated[targetClassIdx].students[targetStudentIdx] = sourceStudent;
      addLog(`[Khối ${activeGrade}] Tráo đổi học sinh ${sourceStudent.HoTen || sourceStudent['Họ và tên']} (Lớp ${updated[sourceClassIdx].className}) với ${updated[targetClassIdx].students[targetStudentIdx].HoTen || updated[targetClassIdx].students[targetStudentIdx]['Họ và tên']} (Lớp ${updated[targetClassIdx].className})`);
    }

    updated[sourceClassIdx].students = sortStudentsAlpha(updated[sourceClassIdx].students);
    updated[targetClassIdx].students = sortStudentsAlpha(updated[targetClassIdx].students);
    updateActiveGradeState({ dividedClasses: updated });
    setSwapModal({ isOpen: false, sourceClassIdx: null, studentIdx: null, student: null, targetClassIdx: 0, targetStudentIdx: -1 });
    setStudentDetailModal({ isOpen: false, classIdx: null, studentIdx: null, student: null });
  };

  const executeDeleteFromClass = () => {
    if (window.confirm("Thầy có chắc chắn muốn LOẠI học sinh này khỏi lớp không? (Học sinh sẽ trở về trạng thái Chưa Xếp Lớp)")) {
      addLog(`[Khối ${activeGrade}] Loại học sinh ${studentDetailModal.student.HoTen || studentDetailModal.student['Họ và tên']} khỏi lớp ${activeState.dividedClasses[studentDetailModal.classIdx].className}`);
      updateActiveGradeState(prev => {
        const updatedClasses = JSON.parse(JSON.stringify(prev.dividedClasses || []));
        updatedClasses[studentDetailModal.classIdx].students.splice(studentDetailModal.studentIdx, 1);
        return { dividedClasses: updatedClasses };
      });
      setStudentDetailModal({ isOpen: false, classIdx: null, studentIdx: null, student: null });
    }
  };

  const executeDeleteStudentCompletely = () => {
    if (window.confirm(`⚠️ CẢNH BÁO NGUY HIỂM:\n\nThầy có chắc chắn muốn XÓA VĨNH VIỄN học sinh "${studentDetailModal.student.HoTen || studentDetailModal.student['Họ và tên']}" khỏi toàn bộ dữ liệu Khối ${activeGrade} không?\n\nHành động này không thể hoàn tác!`)) {
      
      const studentName = studentDetailModal.student.HoTen || studentDetailModal.student['Họ và tên'];
      addLog(`[Khối ${activeGrade}] Xóa vĩnh viễn học sinh "${studentName}" khỏi toàn bộ hệ thống`);
      
      updateActiveGradeState(prev => {
        const updatedClasses = JSON.parse(JSON.stringify(prev.dividedClasses || []));
        updatedClasses[studentDetailModal.classIdx].students.splice(studentDetailModal.studentIdx, 1);
        
        const studentIdToRemove = getStudentId(studentDetailModal.student);
        const updatedData = (prev.data || []).filter(s => getStudentId(s) !== studentIdToRemove);

        return { dividedClasses: updatedClasses, data: updatedData };
      });
      setStudentDetailModal({ isOpen: false, classIdx: null, studentIdx: null, student: null });
    }
  };
  const handleAssignSingleStudent = (student, targetClassIdx) => {
    addLog(`[Khối ${activeGrade}] Xếp học sinh ${student.HoTen || student['Họ và tên']} vào lớp ${activeState.dividedClasses[targetClassIdx].className}`);
    updateActiveGradeState(prev => {
      const updatedClasses = JSON.parse(JSON.stringify(prev.dividedClasses || []));
      updatedClasses[targetClassIdx].students.push(student);
      updatedClasses[targetClassIdx].students = sortStudentsAlpha(updatedClasses[targetClassIdx].students);
      return { dividedClasses: updatedClasses };
    });
    setUnassignedModalOpen(false); setTimeout(() => setUnassignedModalOpen(true), 100);
  };

  const renderA4Page = (cls, studentList, pageIndex, totalPagesForClass, isLastPageOfLastClass = false) => {
    const leftWidth = printConfig.headerLeftWidth || 42;
    const rightWidth = 100 - leftWidth;
    const footWidth = printConfig.footerWidth || 45;
    
    const isShowHeader = printConfig.showHeader !== false; 
    const isShowFooter = printConfig.showFooter !== false; 

    const visibleStudents = (cls?.students || []).filter(s => !s.__isEmptyRow && !(activeState.hiddenPrintIds || []).includes(getStudentId(s)));
    const realStudentCount = visibleStudents.length;
    const nam = visibleStudents.filter(s => (s.GioiTinh || '').toLowerCase() === 'nam').length;
    const nu = visibleStudents.filter(s => (s.GioiTinh || '').toLowerCase() !== 'nam').length;

    const sttWidth = printConfig.colWidths?.STT ? `${printConfig.colWidths.STT}%` : '5%';
    const ghiChuWidth = printConfig.colWidths?.GhiChu ? `${printConfig.colWidths.GhiChu}%` : '18%';

    const cellStyle = {
      border: '0.3pt solid #000', padding: `${printConfig.rowPadding !== undefined ? printConfig.rowPadding : 0.5}px 3px`,
      whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', boxSizing: 'border-box'
    };

    return (
      <div className="a4-page" style={{ 
        fontFamily: '"Times New Roman", Times, serif', width: '210mm', height: '295mm', maxHeight: '295mm', background: '#fff', 
        paddingTop: `${printConfig.marginTop || 5}mm`, paddingBottom: `${printConfig.marginBottom || 5}mm`, paddingLeft: `${printConfig.marginLeft || 5}mm`, paddingRight: `${printConfig.marginRight || 5}mm`,
        margin: '0 auto 20px auto', boxSizing: 'border-box', boxShadow: '0 4px 10px rgba(0,0,0,0.15)', color: '#000', display: 'flex', flexDirection: 'column', position: 'relative', pageBreakInside: 'avoid', breakInside: 'avoid'
      }}>
        <div style={{ marginBottom: `${printConfig.headerSpacing || 2}px` }}>
          {isShowHeader && (
            <div style={{ minHeight: `${printConfig.headerMinHeight || 16}mm`, display: 'flex', justifyContent: 'space-between', marginBottom: '2px' }}>
              <div style={{ textAlign: 'center', width: `${leftWidth}%` }}>
                <div style={{ fontSize: `${printConfig.headerSize}pt`, textTransform: 'uppercase', lineHeight: '1.05' }}>{docFormat.coQuanChuQuan}</div>
                <div style={{ fontSize: `${printConfig.headerSize}pt`, fontWeight: 'bold', textTransform: 'uppercase', lineHeight: '1.05' }}>{docFormat.coQuanBanHanh}</div>
                <hr style={{ width: '45%', border: 'none', borderTop: '0.3pt solid black', margin: '1px auto 0 auto' }} />
              </div>
              <div style={{ textAlign: 'center', width: `${rightWidth}%` }}>
                <div style={{ fontSize: `${printConfig.headerSize}pt`, fontWeight: 'bold', lineHeight: '1.05' }}>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                <div style={{ fontSize: `${printConfig.headerSize}pt`, fontWeight: 'bold', lineHeight: '1.05' }}>Độc lập - Tự do - Hạnh phúc</div>
                <hr style={{ width: '50%', border: 'none', borderTop: '0.3pt solid black', margin: '1px auto 0 auto' }} />
                <div style={{ fontSize: `${printConfig.headerSize}pt`, fontStyle: 'italic', marginTop: '1px', lineHeight: '1.05' }}>{docFormat.diaDanh}</div>
              </div>
            </div>
          )}
          <h2 style={{ textAlign: 'center', fontSize: `${printConfig.headerSize + 2.5}pt`, marginBottom: '1px', marginTop: isShowHeader ? '2px' : '0px', lineHeight: '1.15' }}>{(activeState.docTitle || "DANH SÁCH HỌC SINH").toUpperCase()}</h2>
          <h3 style={{ textAlign: 'center', fontSize: `${printConfig.headerSize + 0.5}pt`, marginTop: '0', marginBottom: '2px', lineHeight: '1.15' }}>
            {(cls?.className || '').toUpperCase()} - {(activeState.academicYear || '').toUpperCase()}
            {totalPagesForClass > 1 && <span style={{ fontSize: '8.5pt', fontWeight: 'normal' }}> (Trang {pageIndex}/{totalPagesForClass})</span>}
          </h3>
          {pageIndex === 1 && (<div style={{ marginBottom: '2px', fontSize: `${printConfig.tableSize - 1}pt`, fontStyle: 'italic', lineHeight: '1.05' }}>Sĩ số: {realStudentCount} học sinh (Nam: {nam}, Nữ: {nu})</div>)}
        </div>

        <div style={{ width: '100%', flex: '1 0 auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: `${printConfig.tableSize || 10.5}pt`, lineHeight: '1.02', tableLayout: 'fixed' }}>
            <thead>
              <tr style={{ pageBreakInside: 'avoid', breakInside: 'avoid' }}>
                <th style={{ ...cellStyle, width: sttWidth, textAlign: 'center' }}>STT</th>
                {(activeState.printColumns || []).map(colId => {
                  const customCol = (activeState.customColumns || []).find(c => c.id === colId);
                  const colTitle = customCol ? customCol.title : colId;
                  const colW = printConfig.colWidths?.[colId];
                  return (<th key={colId} style={{ ...cellStyle, width: colW ? `${colW}%` : 'auto', textAlign: 'center' }}>{String(colTitle || '')}</th>);
                })}
                {showServiceCol && <th style={{ ...cellStyle, width: ghiChuWidth, textAlign: 'center' }}>Ghi chú</th>}
              </tr>
            </thead>
            <tbody>
              {(studentList || []).map((st, sIdx) => {
                if (!st) return null;
                const globalIndex = pageIndex > 1 ? (pageIndex - 1) * (printConfig.fixedRowCount || 54) + sIdx + 1 : sIdx + 1;
                const services = (activeState.priorityList || []).slice(2).filter(col => isYes(st[col])).join(", ");
                return (
                  <tr key={sIdx} style={{ pageBreakInside: 'avoid', breakInside: 'avoid' }}>
                    <td style={{ ...cellStyle, textAlign: 'center' }}>{globalIndex}</td>
                    {(activeState.printColumns || []).map(colId => {
                      const customCol = (activeState.customColumns || []).find(c => c.id === colId);
                      const colTitle = customCol ? customCol.title : colId;
                      const titleStr = String(colTitle || '');
                      const isLeftAlign = titleStr.toLowerCase().includes('tên') || titleStr.toLowerCase().includes('hoten');
                      return (<td key={colId} style={{ ...cellStyle, textAlign: isLeftAlign ? 'left' : 'center' }}>{st[colId] !== undefined ? st[colId] : ""}</td>);
                    })}
                    {showServiceCol && <td style={{ ...cellStyle, fontSize: `${printConfig.tableSize - 1.5}pt`, textAlign: 'center' }}>{services}</td>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {isShowFooter && pageIndex === totalPagesForClass && (
          <div style={{ marginTop: 'auto', width: '100%', display: 'flex', justifyContent: 'flex-end', paddingTop: `${printConfig.footerSpacing || 4}px`, pageBreakInside: 'avoid', breakInside: 'avoid' }}>
            <div style={{ textAlign: 'center', width: `${footWidth}%` }}>
              <div style={{ fontSize: `${printConfig.footerSize}pt`, fontWeight: 'bold', lineHeight: '1.05' }}>{docFormat.chucVu}</div>
              <div style={{ fontSize: `${printConfig.footerSize - 1}pt`, fontStyle: 'italic', marginTop: '1px', lineHeight: '1.05' }}>{docFormat.nguoiKy}</div>
            </div>
          </div>
        )}
      </div>
    );
  };

  const exportStatsToPDF = () => {
    if (classesToExport.length === 0) return alert("Vui lòng chọn ít nhất 1 lớp để xuất báo cáo!");
    
    const printWindow = window.open('', '', 'width=1000,height=800');
    if (!printWindow) return alert("Trình duyệt đã chặn tính năng mở cửa sổ Popup. Vui lòng cho phép popup và thử lại.");
    
    const doc = printWindow.document;
    
    let html = `
      <html>
      <head>
        <title>Bao_Cao_Thong_Ke_Khoi_${activeGrade}</title>
        <style>
          @media print {
            @page { size: A4 landscape; margin: 10mm; }
            body { -webkit-print-color-adjust: exact; }
            .page-break { page-break-after: always; }
          }
          body { font-family: 'Times New Roman', serif; color: #000; margin: 0; padding: 0; }
          .header-table { width: 100%; border: none; margin-bottom: 20px; }
          .header-table td { border: none; text-align: center; vertical-align: top; padding: 0; }
          .title { text-align: center; font-size: 16pt; font-weight: bold; margin: 15px 0 5px 0; }
          .subtitle { text-align: center; font-size: 12pt; font-style: italic; margin-bottom: 20px; }
          .grid-container { 
            display: grid; 
            grid-template-columns: repeat(4, 1fr); 
            gap: 12px; 
            align-items: start;
          }
          .stat-card {
            border: 1pt solid #cbd5e1;
            border-radius: 8px;
            padding: 10px;
            font-size: 11pt;
            box-sizing: border-box;
            page-break-inside: avoid;
            break-inside: avoid;
            background-color: #f8fafc;
          }
          .card-title { font-weight: bold; font-size: 12pt; border-bottom: 1pt dashed #cbd5e1; padding-bottom: 5px; margin-bottom: 8px; color: #0369a1; text-align: center; }
          .stat-row { margin-bottom: 5px; line-height: 1.3; }
          .service-row { margin-top: 8px; border-top: 1pt dashed #cbd5e1; padding-top: 8px; font-size: 10pt; line-height: 1.4; display: flex; flex-wrap: wrap; gap: 8px; color: #0284c7; font-weight: bold; }
          .footer-table { width: 100%; border: none; margin-top: 30px; page-break-inside: avoid; }
          .footer-table td { border: none; text-align: center; vertical-align: top; padding: 0; }
        </style>
      </head>
      <body>
        <table class="header-table">
           <tr>
             <td style="width: 40%; font-size: 12pt;">
               ${docFormat.coQuanChuQuan ? docFormat.coQuanChuQuan.toUpperCase() : ''}<br/>
               <strong>${docFormat.coQuanBanHanh ? docFormat.coQuanBanHanh.toUpperCase() : ''}</strong>
               <hr style="width:40%; border:0.5pt solid black; margin: 4px auto;"/>
             </td>
             <td style="width: 60%; font-size: 12pt;">
               <strong>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</strong><br/>
               <strong>Độc lập - Tự do - Hạnh phúc</strong>
               <hr style="width:30%; border:0.5pt solid black; margin: 4px auto;"/>
             </td>
           </tr>
        </table>

        <div class="title">BÁO CÁO THỐNG KÊ SỐ LƯỢNG HỌC SINH - KHỐI ${activeGrade}</div>
        <div class="subtitle">${activeState.academicYear ? activeState.academicYear.toUpperCase() : ''}</div>

        <div class="grid-container">
    `;

    classesToExport.forEach(cls => {
      const visibleStudents = (cls.students || []).filter(s => !s.__isEmptyRow && !(activeState.hiddenPrintIds || []).includes(getStudentId(s)));
      const nam = visibleStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() === 'nam').length;
      const nu = visibleStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() !== 'nam').length;
      
      const hlT = countAttr(visibleStudents, 'HL', 'T'); const hlK = countAttr(visibleStudents, 'HL', 'K'); const hlD = countAttr(visibleStudents, 'HL', 'Đ'); const hlCD = countAttr(visibleStudents, 'HL', 'CĐ');
      const hkT = countAttr(visibleStudents, 'HK', 'T'); const hkK = countAttr(visibleStudents, 'HK', 'K'); const hkD = countAttr(visibleStudents, 'HK', 'Đ');

      const formatPrintStat = (text, isHL) => {
         const color = isHL ? '#dc2626' : '#1d4ed8';
         const parts = text.split(' - ');
         return parts.map(p => {
            const m = p.match(/([A-ZĐ]+)\((\d+)\)/);
            if(m) return `${m[1]}(<strong style="color: ${color};">${m[2]}</strong>)`;
            return p;
         }).join(' - ');
      };

      html += `
        <div class="stat-card">
          <div class="card-title">${cls.className} (Sĩ số: ${nam + nu})</div>
          <div class="stat-row">👨 Nam: ${nam} &nbsp;|&nbsp; 👩 Nữ: ${nu}</div>
          <div class="stat-row" style="color: #334155;"><strong>HL:</strong> ${formatPrintStat(`T(${hlT}) - K(${hlK}) - Đ(${hlD}) - CĐ(${hlCD})`, true)}</div>
          <div class="stat-row" style="color: #334155;"><strong>HK:</strong> ${formatPrintStat(`T(${hkT}) - K(${hkK}) - Đ(${hkD})`, false)}</div>
          <div class="service-row">
            <span>❄️ Lạnh: ${visibleStudents.filter(hasMayLanh).length}</span>
            <span>🧪 STEM: ${visibleStudents.filter(hasStem).length}</span>
            <span>🗣️ AVBN: ${visibleStudents.filter(hasAVBN).length}</span>
            <span>🧘 KNS: ${visibleStudents.filter(hasKNS).length}</span>
            <span>💻 IC3: ${visibleStudents.filter(hasIC3).length}</span>
          </div>
        </div>
      `;
    });

    html += `
        </div>
        <table class="footer-table">
          <tr>
            <td style="width: 50%;"></td>
            <td style="width: 50%; font-size: 12pt;">
              <i>${docFormat.diaDanh || ''}</i><br/><br/>
              <strong>${docFormat.chucVu || ''}</strong><br/><br/><br/><br/>
              <i>${docFormat.nguoiKy || ''}</i>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    doc.open();
    doc.write(html);
    doc.close();

    setTimeout(() => {
      printWindow.focus();
      printWindow.print();
      printWindow.close();
    }, 500);
  };

  const renderModals = () => (
    <>
      {showLogModal && (
        <div className="no-print" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1300 }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '12px', width: '700px', maxWidth: '95%', maxHeight: '85vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', borderBottom: '2px solid #e2e8f0', paddingBottom: '10px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#1e40af', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <History size={22} color="#2563eb" /> Nhật Ký Hệ Thống (Audit Trail)
              </h3>
              <button onClick={() => setShowLogModal(false)} style={{ border: 'none', background: '#f1f5f9', padding: '6px', borderRadius: '50%', cursor: 'pointer' }}><X size={20} color="#64748b" /></button>
            </div>
            
            <div style={{ flex: 1, overflowY: 'auto', paddingRight: '5px' }}>
              {activityLogs.length === 0 ? (
                <div style={{ textAlign: 'center', padding: '40px 20px', color: '#64748b' }}>Chưa có lịch sử thay đổi nào được ghi nhận.</div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  {activityLogs.map((log) => (
                    <div key={log.id} style={{ backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderLeft: '4px solid #3b82f6', borderRadius: '6px', padding: '12px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px', fontSize: '12px', color: '#64748b' }}>
                        <span style={{ fontWeight: 'bold', color: '#0f172a' }}>{log.time}</span>
                        <span>Thực hiện bởi: {log.user}</span>
                      </div>
                      <div style={{ fontSize: '14px', color: '#334155', fontWeight: 'bold', lineHeight: '1.4' }}>
                        {log.action}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
            
            <div style={{ marginTop: '15px', textAlign: 'right', borderTop: '1px solid #e2e8f0', paddingTop: '10px' }}>
              <button onClick={() => { if(window.confirm('Xóa toàn bộ nhật ký?')) setActivityLogs([]); localStorage.removeItem('app_activity_logs'); }} style={{ padding: '8px 15px', backgroundColor: '#fee2e2', color: '#ef4444', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
                Xóa lịch sử
              </button>
            </div>
          </div>
        </div>
      )}

      {showGhiChuModal && (
        <div className="no-print" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '12px', width: '850px', maxWidth: '95%', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', borderBottom: '2px solid #e2e8f0', paddingBottom: '10px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#9d174d', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <FileText size={22} color="#ec4899" /> Thống Kê Học Sinh Có Ghi Chú (Khối {activeGrade})
              </h3>
              <button onClick={() => setShowGhiChuModal(false)} style={{ border: 'none', background: '#fce7f3', padding: '6px', borderRadius: '50%', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <X size={20} color="#be185d" />
              </button>
            </div>
            
            <div style={{ display: 'flex', gap: '10px', marginBottom: '15px', flexWrap: 'wrap' }}>
              <input
                type="text"
                placeholder="Lọc theo nội dung ghi chú..."
                value={noteSearchQuery}
                onChange={(e) => setNoteSearchQuery(e.target.value)}
                style={{ flex: 1, minWidth: '200px', padding: '8px 12px', border: '1px solid #f9a8d4', borderRadius: '6px', fontSize: '13px', outline: 'none' }}
              />
              <select
                value={noteSortOption}
                onChange={(e) => setNoteSortOption(e.target.value)}
                style={{ padding: '8px 12px', border: '1px solid #f9a8d4', borderRadius: '6px', fontSize: '13px', outline: 'none', backgroundColor: '#fff', cursor: 'pointer', fontWeight: 'bold', color: '#831843' }}
              >
                <option value="NAME_ASC">Sắp xếp: Tên (A-Z)</option>
                <option value="NAME_DESC">Sắp xếp: Tên (Z-A)</option>
                <option value="NOTE_ASC">Sắp xếp: Nội dung Ghi chú (A-Z)</option>
                <option value="CLASS_ASC">Sắp xếp: Lớp (A-Z)</option>
              </select>
            </div>

            <div style={{ maxHeight: '60vh', overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: '8px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead style={{ position: 'sticky', top: 0, backgroundColor: '#fbcfe8', zIndex: 1 }}>
                  <tr>
                    <th style={{ padding: '10px', borderBottom: '2px solid #f9a8d4', textAlign: 'center', width: '50px', color: '#831843' }}>STT</th>
                    <th style={{ padding: '10px', borderBottom: '2px solid #f9a8d4', textAlign: 'left', color: '#831843' }}>Họ và Tên</th>
                    <th style={{ padding: '10px', borderBottom: '2px solid #f9a8d4', textAlign: 'center', color: '#831843' }}>Lớp Hiện Tại</th>
                    <th style={{ padding: '10px', borderBottom: '2px solid #f9a8d4', textAlign: 'left', color: '#831843' }}>Nội dung Ghi Chú</th>
                  </tr>
                </thead>
                <tbody>
                  {(() => {
                    let rawNotesList = [];
                    
                    (activeState.dividedClasses || []).forEach(cls => {
                      (cls.students || []).forEach(st => {
                        const note = getGhiChu(st);
                        if (note) {
                          rawNotesList.push({
                            st,
                            className: cls.className,
                            note,
                            isAssigned: true
                          });
                        }
                      });
                    });
                    
                    unassignedStudents.forEach(st => {
                      const note = getGhiChu(st);
                      if (note) {
                        rawNotesList.push({
                          st,
                          className: 'Chưa xếp lớp',
                          note,
                          isAssigned: false
                        });
                      }
                    });

                    if (noteSearchQuery.trim()) {
                      const query = noteSearchQuery.toLowerCase();
                      rawNotesList = rawNotesList.filter(item => item.note.toString().toLowerCase().includes(query));
                    }

                    rawNotesList.sort((a, b) => {
                      if (noteSortOption === 'NAME_ASC' || noteSortOption === 'NAME_DESC') {
                        const getParts = (name) => {
                          if (!name) return { ten: '', hoDem: '' };
                          const p = name.toString().trim().split(/\s+/);
                          const ten = p.pop() || '';
                          const hoDem = p.join(' ');
                          return { ten, hoDem };
                        };
                        const aParts = getParts(a.st.HoTen || a.st['Họ và tên']);
                        const bParts = getParts(b.st.HoTen || b.st['Họ và tên']);
                        let cmp = aParts.ten.localeCompare(bParts.ten, 'vi');
                        if (cmp === 0) cmp = aParts.hoDem.localeCompare(bParts.hoDem, 'vi');
                        return noteSortOption === 'NAME_ASC' ? cmp : -cmp;
                      } else if (noteSortOption === 'NOTE_ASC') {
                        return a.note.toString().localeCompare(b.note.toString(), 'vi');
                      } else if (noteSortOption === 'CLASS_ASC') {
                        return a.className.localeCompare(b.className, 'vi');
                      }
                      return 0;
                    });

                    if (rawNotesList.length === 0) {
                      return <tr><td colSpan="4" style={{ padding: '20px', textAlign: 'center', color: '#64748b' }}>Không tìm thấy học sinh nào khớp với điều kiện lọc trong Khối {activeGrade}.</td></tr>;
                    }

                    return rawNotesList.map((item, idx) => (
                      <tr key={getStudentId(item.st)} style={{ backgroundColor: idx % 2 === 0 ? '#fdf2f8' : '#fff' }}>
                        <td style={{ padding: '8px', borderBottom: '1px solid #fce7f3', textAlign: 'center' }}>{idx + 1}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #fce7f3', fontWeight: 'bold', color: '#334155' }}>{item.st.HoTen || item.st['Họ và tên']} <span style={{fontWeight: 'normal', color: '#64748b', fontSize: '11px'}}>({getDob(item.st)})</span></td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #fce7f3', textAlign: 'center', fontWeight: 'bold', color: item.isAssigned ? '#0284c7' : '#b45309' }}>{item.className}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #fce7f3', color: '#ef4444', fontWeight: 'bold', fontStyle: 'italic' }}>{item.note}</td>
                      </tr>
                    ));
                  })()}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {showStatsModal && (
        <div className="no-print" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '12px', width: '900px', maxWidth: '95%', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', borderBottom: '2px solid #e2e8f0', paddingBottom: '10px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#0f172a', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <BarChart2 size={22} color="#ec4899" /> Báo Cáo Thống Kê Số Lượng ({classesToExport.length} Lớp đang chọn in)
              </h3>
              <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                <input type="file" ref={gvcnInputRef} accept=".xlsx, .xls" onChange={handleImportGVCN} style={{ display: 'none' }} />
                <button onClick={() => gvcnInputRef.current.click()} style={{ padding: '6px 12px', backgroundColor: '#8b5cf6', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }} title="Nạp danh sách Giáo viên chủ nhiệm từ Excel">
                  <Upload size={16} /> Nạp File GVCN
                </button>
                <button onClick={exportStatsToExcel} style={{ padding: '6px 12px', backgroundColor: '#16a34a', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
                  <Download size={16} /> Xuất Excel Thống Kê
                </button>
                <button onClick={exportStatsToPDF} style={{ padding: '6px 12px', backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
                  <Printer size={16} /> Xuất PDF Báo Cáo
                </button>
                <button onClick={() => setShowStatsModal(false)} style={{ border: 'none', background: '#f1f5f9', padding: '6px', borderRadius: '50%', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                  <X size={20} color="#64748b" />
                </button>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(350px, 1fr))', gap: '15px' }}>
              {classesToExport.map((cls, idx) => {
                const visibleStudents = (cls.students || []).filter(s => !s.__isEmptyRow && !(activeState.hiddenPrintIds || []).includes(getStudentId(s)));
                const nam = visibleStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() === 'nam').length;
                const nu = visibleStudents.filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() !== 'nam').length;
                const hlT = countAttr(visibleStudents, 'HL', 'T'); const hlK = countAttr(visibleStudents, 'HL', 'K'); const hlD = countAttr(visibleStudents, 'HL', 'Đ'); const hlCD = countAttr(visibleStudents, 'HL', 'CĐ');
                const hkT = countAttr(visibleStudents, 'HK', 'T'); const hkK = countAttr(visibleStudents, 'HK', 'K'); const hkD = countAttr(visibleStudents, 'HK', 'Đ');
                
                return (
                  <div key={idx} style={{ padding: '15px', backgroundColor: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: '10px', fontSize: '15px', color: '#334155', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                    <div style={{ fontWeight: 'bold', color: '#0369a1', fontSize: '17px', marginBottom: '8px', borderBottom: '1px dashed #cbd5e1', paddingBottom: '6px' }}>{cls.className} (Sĩ số: {nam + nu})</div>
                    <div style={{ fontWeight: 'bold', marginBottom: '6px', color: '#0f172a' }}>👨 Nam: {nam} &nbsp;&nbsp;|&nbsp;&nbsp; 👩 Nữ: {nu}</div>
                    <div style={{ marginBottom: '4px' }}><strong>HL:</strong> {formatStatText(`T(${hlT}) - K(${hlK}) - Đ(${hlD}) - CĐ(${hlCD})`, 'HL')}</div>
                    <div style={{ marginBottom: '10px' }}><strong>HK:</strong> {formatStatText(`T(${hkT}) - K(${hkK}) - Đ(${hkD})`, 'HK')}</div>
                    <div style={{ display: 'flex', gap: '12px', flexWrap: 'nowrap', whiteSpace: 'nowrap', fontWeight: 'bold', color: '#0284c7', overflowX: 'auto', paddingBottom: '4px', backgroundColor: '#e0f2fe', padding: '8px', borderRadius: '6px' }}>
                      <span>❄️ Lạnh: {visibleStudents.filter(hasMayLanh).length}</span>
                      <span>🧪 STEM: {visibleStudents.filter(hasStem).length}</span>
                      <span>🗣️ AVBN: {visibleStudents.filter(hasAVBN).length}</span>
                      <span>🧘 KNS: {visibleStudents.filter(hasKNS).length}</span>
                      <span>💻 IC3: {visibleStudents.filter(hasIC3).length}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {printHideModalOpen && (
        <div className="no-print" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '12px', width: '700px', maxWidth: '95%', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', borderBottom: '2px solid #e2e8f0', paddingBottom: '10px' }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '18px', color: '#0f172a', display: 'flex', alignItems: 'center', gap: '8px' }}><EyeOff size={22} color="#64748b" /> Quản lý Ẩn/Hiện Học Sinh Khi In</h3>
                <div style={{ fontSize: '13px', color: '#ef4444', marginTop: '4px' }}>*Học sinh bị "tắt mắt" sẽ không xuất hiện trong bản in PDF/Excel và không được cộng vào sĩ số tổng.</div>
              </div>
              <button onClick={() => setPrintHideModalOpen(false)} style={{ border: 'none', background: '#f1f5f9', padding: '6px', borderRadius: '50%', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><X size={20} color="#64748b" /></button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
              {classesToExport.map((cls, idx) => (
                <div key={idx} style={{ border: '1px solid #cbd5e1', borderRadius: '8px', overflow: 'hidden' }}>
                  <div style={{ backgroundColor: '#f1f5f9', padding: '8px 12px', fontWeight: 'bold', color: '#0369a1', borderBottom: '1px solid #cbd5e1' }}>{cls.className}</div>
                  <div style={{ maxHeight: '200px', overflowY: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                      <tbody>
                        {(cls.students || []).map((st, sIdx) => {
                          const stId = getStudentId(st);
                          const isHidden = (activeState.hiddenPrintIds || []).includes(stId);
                          return (
                            <tr key={sIdx} style={{ borderBottom: '1px solid #f1f5f9', backgroundColor: isHidden ? '#fef2f2' : '#fff' }}>
                              <td style={{ padding: '6px 12px', color: isHidden ? '#ef4444' : '#334155', textDecoration: isHidden ? 'line-through' : 'none' }}>
                                {sIdx + 1}. {st.HoTen || st['Họ và tên']} ({getDob(st)})
                              </td>
                              <td style={{ padding: '6px 12px', textAlign: 'right' }}>
                                <button
                                  onClick={() => {
                                    updateActiveGradeState(prev => {
                                      let currentHidden = prev.hiddenPrintIds || [];
                                      if (isHidden) currentHidden = currentHidden.filter(id => id !== stId);
                                      else currentHidden = [...currentHidden, stId];
                                      return { hiddenPrintIds: currentHidden };
                                    });
                                  }}
                                  style={{ padding: '4px 8px', borderRadius: '4px', border: 'none', cursor: 'pointer', backgroundColor: isHidden ? '#ef4444' : '#10b981', color: '#fff', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px', marginLeft: 'auto' }}
                                >
                                  {isHidden ? <><EyeOff size={14} /> Đang Ẩn</> : <><Eye size={14} /> Hiển thị</>}
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {unassignedModalOpen && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '12px', width: '800px', maxWidth: '95%', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', borderBottom: '2px solid #e2e8f0', paddingBottom: '10px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#b45309', display: 'flex', alignItems: 'center', gap: '8px' }}><UserPlus size={22} color="#f59e0b" /> Danh sách Học Sinh Chưa Xếp Lớp ({unassignedStudents.length} HS)</h3>
              <button onClick={() => setUnassignedModalOpen(false)} style={{ border: 'none', background: '#f1f5f9', padding: '6px', borderRadius: '50%', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><X size={20} color="#64748b" /></button>
            </div>

            {unassignedStudents.length === 0 ? (
              <div style={{ padding: '20px', textAlign: 'center', color: '#166534', backgroundColor: '#dcfce7', borderRadius: '8px', fontWeight: 'bold' }}>Tất cả học sinh đều đã được xếp lớp!</div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr style={{ backgroundColor: '#fef3c7', color: '#92400e' }}>
                    <th style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'left' }}>STT</th>
                    <th style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'left' }}>Họ Tên</th>
                    <th style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'left' }}>Ngày Sinh</th>
                    <th style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center' }}>HL</th>
                    <th style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center' }}>HK</th>
                    <th style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center' }}>Lớp Cũ</th>
                    <th style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center' }}>Hành động: Chọn Lớp Mới</th>
                  </tr>
                </thead>
                <tbody>
                  {unassignedStudents.map((st, idx) => (
                    <tr key={idx} style={{ backgroundColor: idx % 2 === 0 ? '#fff' : '#fffbeb' }}>
                      <td style={{ padding: '8px', border: '1px solid #fde68a' }}>{idx + 1}</td>
                      <td style={{ padding: '8px', border: '1px solid #fde68a', fontWeight: 'bold', color: '#b45309' }}>{st.HoTen || st['Họ và tên']}</td>
                      <td style={{ padding: '8px', border: '1px solid #fde68a' }}>{getDob(st)}</td>
                      <td style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center', fontWeight: 'bold', color: '#dc2626' }}>{st.HL || ''}</td>
                      <td style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center', fontWeight: 'bold', color: '#1d4ed8' }}>{st.HK || ''}</td>
                      <td style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center', fontWeight: 'bold' }}>{st.LopCu || st['Lớp Cũ'] || 'Không rõ'}</td>
                      <td style={{ padding: '8px', border: '1px solid #fde68a', textAlign: 'center' }}>
                        <div style={{ display: 'flex', gap: '6px', justifyContent: 'center' }}>
                          <select id={`select_class_${idx}`} style={{ padding: '4px', borderRadius: '4px', border: '1px solid #fcd34d', fontWeight: 'bold' }}>
                            <option value="">-- Chọn Lớp --</option>
                            {(activeState.dividedClasses || []).map((cls, cIdx) => (<option key={cIdx} value={cIdx}>{cls.className} (Đang có {cls.students.length} HS)</option>))}
                          </select>
                          <button 
                            onClick={() => { const sel = document.getElementById(`select_class_${idx}`); if(sel.value !== "") { handleAssignSingleStudent(st, parseInt(sel.value)); } else { alert("Vui lòng chọn lớp để đưa học sinh vào!"); } }}
                            style={{ backgroundColor: '#f59e0b', color: '#fff', border: 'none', borderRadius: '4px', padding: '4px 8px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}
                          >Đưa vào</button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {studentDetailModal.isOpen && studentDetailModal.student && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1100 }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '12px', width: '450px', maxWidth: '90%', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', borderBottom: '2px solid #e2e8f0', paddingBottom: '10px' }}>
              <h3 style={{ margin: 0, color: '#0369a1', fontSize: '18px' }}>Quản Lý Cá Nhân Học Sinh</h3>
              <button onClick={() => setStudentDetailModal({ isOpen: false, classIdx: null, studentIdx: null, student: null })} style={{ border: 'none', background: 'none', cursor: 'pointer' }}><X size={20} color="#64748b" /></button>
            </div>

            <div style={{ backgroundColor: '#f0f9ff', padding: '15px', borderRadius: '8px', border: '1px solid #bae6fd', marginBottom: '15px' }}>
              <div style={{ fontSize: '16px', fontWeight: 'bold', color: '#0284c7', marginBottom: '6px' }}>{studentDetailModal.student.HoTen || studentDetailModal.student['Họ và tên']} <span style={{fontSize: '14px', color: '#64748b', fontWeight: 'normal'}}>({getDob(studentDetailModal.student)})</span></div>
              <div style={{ display: 'flex', gap: '15px', fontSize: '13px', color: '#334155', marginBottom: '8px' }}>
                <span><strong>Giới tính:</strong> {studentDetailModal.student.GioiTinh}</span>
                <span><strong style={{color: '#dc2626'}}>HL:</strong> {studentDetailModal.student.HL}</span>
                <span><strong style={{color: '#1d4ed8'}}>HK:</strong> {studentDetailModal.student.HK}</span>
              </div>
              <div style={{ fontSize: '13px', color: '#334155', borderTop: '1px dashed #93c5fd', paddingTop: '8px' }}>
                <strong>Dịch vụ đăng ký:</strong> 
                <span style={{ color: '#0369a1', fontWeight: 'bold', marginLeft: '6px' }}>
                  {hasMayLanh(studentDetailModal.student) && "❄️ Máy lạnh, "}
                  {hasStem(studentDetailModal.student) && "🧪 STEM, "}
                  {hasAVBN(studentDetailModal.student) && "🗣️ AVBN, "}
                  {hasKNS(studentDetailModal.student) && "🧘 KNS, "}
                  {hasIC3(studentDetailModal.student) && "💻 IC3"}
                  {(!hasMayLanh(studentDetailModal.student) && !hasStem(studentDetailModal.student) && !hasAVBN(studentDetailModal.student) && !hasKNS(studentDetailModal.student) && !hasIC3(studentDetailModal.student)) && "Không có"}
                </span>
              </div>
              {getGhiChu(studentDetailModal.student) && (
                <div style={{ fontSize: '13px', color: '#334155', marginTop: '6px' }}>
                  <strong>Ghi chú đặc biệt:</strong> <span style={{ color: '#ef4444', fontStyle: 'italic', fontWeight: 'bold' }}>{getGhiChu(studentDetailModal.student)}</span>
                </div>
              )}
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <button onClick={() => { openSwapModal(studentDetailModal.classIdx, studentDetailModal.studentIdx, studentDetailModal.student); }} style={{ padding: '10px', backgroundColor: '#0284c7', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', fontSize: '14px' }}>
                <ArrowRightLeft size={18} /> Mở Bảng Chuyển Lớp / Tráo Đổi
              </button>
              <button onClick={executeDeleteFromClass} style={{ padding: '10px', backgroundColor: '#fff', color: '#f59e0b', border: '1px solid #fcd34d', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', fontSize: '14px' }}>
                <Layers size={18} /> Loại khỏi lớp này (Về danh sách chưa xếp)
              </button>
              <button onClick={executeDeleteStudentCompletely} style={{ padding: '10px', backgroundColor: '#fff', color: '#ef4444', border: '1px solid #fecaca', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', fontSize: '14px' }}>
                <Trash2 size={18} /> Xóa vĩnh viễn khỏi hệ thống
              </button>
            </div>
          </div>
        </div>
      )}

      {swapModal.isOpen && activeState.dividedClasses && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1200 }}>
          <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '10px', width: '450px', maxWidth: '90%', boxShadow: '0 4px 12px rgba(0,0,0,0.15)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', borderBottom: '1px solid #e2e8f0', paddingBottom: '10px' }}>
              <h3 style={{ margin: 0, color: '#0f172a', fontSize: '16px' }}>⇄ Chuyển / Tráo Học Sinh (Khối {activeGrade})</h3>
              <button onClick={() => setSwapModal({ ...swapModal, isOpen: false })} style={{ border: 'none', background: 'none', cursor: 'pointer' }}><X size={18} /></button>
            </div>

            <div style={{ fontSize: '13px', marginBottom: '12px' }}>
              <strong>Học sinh chuyển:</strong> <span style={{ color: '#0284c7', fontWeight: 'bold' }}>{swapModal.student?.HoTen || swapModal.student?.['Họ và tên']}</span><br/>
              <strong>Lớp hiện tại đang học:</strong> <span style={{ color: '#ef4444', fontWeight: 'bold' }}>{activeState.dividedClasses[swapModal.sourceClassIdx]?.className}</span>
            </div>

            <div style={{ marginBottom: '12px' }}>
              <label style={{ fontWeight: 'bold', fontSize: '13px' }}>Chuyển sang Lớp:</label>
              <select value={swapModal.targetClassIdx} onChange={(e) => setSwapModal({ ...swapModal, targetClassIdx: parseInt(e.target.value), targetStudentIdx: -1 })} style={{ width: '100%', padding: '8px', marginTop: '4px', borderRadius: '6px', border: '1px solid #94a3b8', fontWeight: 'bold' }}>
                {activeState.dividedClasses.map((cls, idx) => (<option key={idx} value={idx} disabled={idx === swapModal.sourceClassIdx}>{cls.className} (Đang có {(cls.students || []).length}/{activeState.maxStudentsPerClass || 54} HS)</option>))}
              </select>
            </div>

            <div style={{ marginBottom: '15px' }}>
              <label style={{ fontWeight: 'bold', fontSize: '13px' }}>Hành động:</label>
              <select value={swapModal.targetStudentIdx} onChange={(e) => setSwapModal({ ...swapModal, targetStudentIdx: parseInt(e.target.value) })} style={{ width: '100%', padding: '8px', marginTop: '4px', borderRadius: '6px', border: '1px solid #94a3b8' }}>
                <option value={-1}>Chuyển thẳng học sinh qua (Tăng sĩ số lớp kia lên 1)</option>
                <optgroup label="Hoặc chọn 1 học sinh ở Lớp đích để TRÁO ĐỔI 1-1:">
                  {activeState.dividedClasses[swapModal.targetClassIdx]?.students?.map((st, idx) => (<option key={idx} value={idx}>Tráo đổi với: {st.HoTen || st['Họ và tên']} ({st.GioiTinh} | HL:{st.HL})</option>))}
                </optgroup>
              </select>
            </div>

            <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
              <button onClick={() => setSwapModal({ ...swapModal, isOpen: false })} style={{ padding: '8px 14px', backgroundColor: '#e2e8f0', color: '#475569', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>Hủy</button>
              <button onClick={executeManualSwap} style={{ padding: '8px 14px', backgroundColor: '#0284c7', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}><UserCheck size={16} /> Xác Nhận Chuyển / Tráo</button>
            </div>
          </div>
        </div>
      )}
    </>
  );

  if (currentView === 'PRINT') {
    return (
      <div style={{ backgroundColor: '#cbd5e1', minHeight: '100vh', paddingBottom: '50px' }}>
        <style>{`
          @media print {
            .no-print { display: none !important; }
            html, body { background-color: white !important; margin: 0 !important; padding: 0 !important; width: auto !important; height: auto !important; overflow: visible !important; }
            @page { size: A4 portrait; margin: 0mm !important; }
            .print-content-container { margin-top: 0 !important; padding-top: 0 !important; }
            .class-print-wrapper { display: contents !important; }
            .a4-page { box-shadow: none !important; margin: 0 !important; width: 210mm !important; height: 295mm !important; max-height: 295mm !important; box-sizing: border-box !important; overflow: visible !important; display: flex !important; flex-direction: column !important; position: relative !important; }
            tr { page-break-inside: avoid !important; break-inside: avoid !important; }
            th, td { border-width: 0.3pt !important; }
          }
        `}</style>

        {/* Thanh công cụ IN */}
        <div className="no-print" style={{ backgroundColor: '#fff', padding: '12px 20px', display: 'flex', flexWrap: 'wrap', gap: '12px', alignItems: 'flex-start', boxShadow: '0 2px 4px rgba(0,0,0,0.1)', position: 'sticky', top: 0, zIndex: 10 }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <button onClick={() => setCurrentView('MAIN')} style={{ padding: '8px 12px', backgroundColor: '#e2e8f0', color: '#0f172a', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><Home size={16} /> Về trang Quản lý</button>
            <div style={{ display: 'flex', gap: '8px' }}>
              <div style={{ position: 'relative' }}>
                <button onClick={() => setShowColMenu(!showColMenu)} style={{ width: '100%', padding: '8px 12px', backgroundColor: '#f59e0b', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><Columns size={16} /> Quản lý Cột In</button>
                {showColMenu && (
                  <div style={{ position: 'absolute', top: '100%', left: 0, marginTop: '8px', backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '15px', width: '380px', boxShadow: '0 10px 25px rgba(0,0,0,0.2)', zIndex: 100 }}>
                    <button onClick={addCustomColumn} style={{ width: '100%', padding: '8px', backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px', marginBottom: '12px' }}><PlusCircle size={16} /> + Thêm Cột Tùy Chỉnh Mới</button>
                    <h4 style={{ margin: '0 0 8px 0', fontSize: '13px', color: '#0f172a' }}>1. Tùy chỉnh Độ rộng Cột (%):</h4>
                    <div style={{ maxHeight: '200px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '10px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#f0fdf4', padding: '4px 6px', borderRadius: '4px', border: '1px solid #bbf7d0' }}>
                        <span style={{ flex: 1, fontSize: '12px', fontWeight: 'bold', color: '#166534' }}>Cột: STT</span>
                        <input type="number" min="1" max="100" value={printConfig.colWidths?.STT || ''} onChange={e => handleColWidthChange('STT', e.target.value)} placeholder="%" style={{ width: '50px', padding: '4px', fontSize: '12px', border: '1px solid #86efac', borderRadius: '4px', textAlign: 'center' }} />
                      </div>
                      {showServiceCol && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#f0fdf4', padding: '4px 6px', borderRadius: '4px', border: '1px solid #bbf7d0' }}>
                          <span style={{ flex: 1, fontSize: '12px', fontWeight: 'bold', color: '#166534' }}>Cột: Ghi chú</span>
                          <input type="number" min="1" max="100" value={printConfig.colWidths?.GhiChu || ''} onChange={e => handleColWidthChange('GhiChu', e.target.value)} placeholder="%" style={{ width: '50px', padding: '4px', fontSize: '12px', border: '1px solid #86efac', borderRadius: '4px', textAlign: 'center' }} />
                        </div>
                      )}
                      {[...(activeState.excelColumns || []), ...(activeState.customColumns || []).map(c => c.id)].map(colItem => {
                        const isCustom = colItem.startsWith('custom_'); const customObj = isCustom ? (activeState.customColumns || []).find(c => c.id === colItem) : null;
                        const colTitle = isCustom ? customObj.title : colItem; const isChecked = (activeState.printColumns || []).includes(colItem);
                        return (
                          <div key={colItem} style={{ display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: isChecked ? '#eff6ff' : '#f8fafc', padding: '4px 6px', borderRadius: '4px', border: isChecked ? '1px solid #bfdbfe' : '1px solid #e2e8f0' }}>
                            <input type="checkbox" checked={isChecked} onChange={(e) => { if (e.target.checked) updateActiveGradeState(p => ({ printColumns: [...(p.printColumns || []), colItem] })); else updateActiveGradeState(p => ({ printColumns: (p.printColumns || []).filter(c => c !== colItem) })); }} />
                            {isCustom ? (<input type="text" value={colTitle} onChange={(e) => updateCustomColumnTitle(colItem, e.target.value)} style={{ flex: 1, padding: '4px', fontSize: '12px', border: '1px solid #cbd5e1', borderRadius: '4px', fontWeight: 'bold' }} />) : (<span style={{ flex: 1, fontSize: '12px', color: isChecked ? '#1e3a8a' : '#64748b', fontWeight: isChecked ? 'bold' : 'normal' }}>{colTitle}</span>)}
                            {isChecked && (<input type="number" min="1" max="100" value={printConfig.colWidths?.[colItem] || ''} onChange={e => handleColWidthChange(colItem, e.target.value)} placeholder="%" title="Độ rộng cột (%)" style={{ width: '50px', padding: '4px', fontSize: '12px', border: '1px solid #93c5fd', borderRadius: '4px', textAlign: 'center' }} />)}
                            {isCustom && (<button onClick={() => removeCustomColumn(colItem)} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#ef4444' }} title="Xóa cột"><Trash2 size={14} /></button>)}
                          </div>
                        )
                      })}
                    </div>
                    <hr style={{ margin: '10px 0', borderTop: '1px solid #e2e8f0' }}/><label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', cursor: 'pointer', fontWeight: 'bold', color: '#0369a1' }}><input type="checkbox" checked={showServiceCol} onChange={e => setShowServiceCol(e.target.checked)} /> Tự tạo cột Ghi chú (Dịch vụ)</label>
                  </div>
                )}
              </div>
              <button onClick={() => setPrintHideModalOpen(true)} style={{ padding: '8px 12px', backgroundColor: '#94a3b8', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><EyeOff size={16} /> Quản lý Ẩn/Hiện HS</button>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '6px', borderLeft: '2px solid #e2e8f0', borderRight: '2px solid #e2e8f0', paddingLeft: '8px', paddingRight: '8px' }}>
            {['6', '7', '8', '9'].map(gKey => (
              <button key={gKey} onClick={() => setActiveGrade(gKey)} style={{ padding: '6px 12px', backgroundColor: activeGrade === gKey ? '#0284c7' : '#f1f5f9', color: activeGrade === gKey ? '#fff' : '#334155', border: 'none', borderRadius: '6px', fontWeight: 'bold', fontSize: '12px', cursor: 'pointer' }}>Khối {gKey}</button>
            ))}
          </div>

          <button onClick={handleAutoFitA4} style={{ padding: '8px 14px', backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', boxShadow: '0 2px 4px rgba(16, 185, 129, 0.3)' }}><Sparkles size={16} /> Gom khít chuẩn 11pt</button>
          <button onClick={saveGlobalConfig} style={{ padding: '8px 14px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', boxShadow: '0 2px 4px rgba(37, 99, 235, 0.3)' }} title="Lưu các thông tin văn bản, lề in, chức danh vào máy để dùng tự động cho các lần sau"><Save size={16} /> Lưu Cấu Hình</button>
          
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center', backgroundColor: '#f8fafc', padding: '6px 10px', borderRadius: '6px', border: '1px solid #e2e8f0', flexWrap: 'wrap' }}>
            <Settings size={16} color="#64748b"/>
            <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#475569' }}>Size Tiêu đề</label><input type="number" min="8" max="40" step="1" value={printConfig.headerSize} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, headerSize: Number(e.target.value) } }))} style={{ width: '45px', border: '1px solid #cbd5e1', borderRadius: '4px', textAlign: 'center', fontSize: '12px' }} /></div>
            <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#0284c7' }}>Size Bảng</label><input type="number" min="8" max="16" step="0.5" value={printConfig.tableSize} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, tableSize: Number(e.target.value) } }))} style={{ width: '45px', border: '2px solid #0284c7', borderRadius: '4px', textAlign: 'center', fontSize: '12px', fontWeight: 'bold', color: '#0284c7' }} /></div>
            <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#0284c7' }}>Giãn dòng</label><input type="number" min="0" max="20" step="0.1" value={printConfig.rowPadding !== undefined ? printConfig.rowPadding : 0.5} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, rowPadding: Number(e.target.value) } }))} style={{ width: '45px', border: '2px solid #0284c7', borderRadius: '4px', textAlign: 'center', fontSize: '12px', fontWeight: 'bold', color: '#0284c7' }} /></div>
            <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#475569' }}>Size Chữ ký</label><input type="number" min="8" max="30" step="1" value={printConfig.footerSize} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, footerSize: Number(e.target.value) } }))} style={{ width: '45px', border: '1px solid #cbd5e1', borderRadius: '4px', textAlign: 'center', fontSize: '12px' }} /></div>
            <div style={{ borderLeft: '1px solid #cbd5e1', paddingLeft: '8px', display: 'flex', gap: '4px' }}>
              <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#475569' }}>Lề Trái</label><input type="number" min="0" max="40" value={printConfig.marginLeft} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, marginLeft: Number(e.target.value) } }))} style={{ width: '35px', border: '1px solid #cbd5e1', borderRadius: '4px', textAlign: 'center', fontSize: '12px' }} /></div>
              <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#475569' }}>Lề Phải</label><input type="number" min="0" max="40" value={printConfig.marginRight} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, marginRight: Number(e.target.value) } }))} style={{ width: '35px', border: '1px solid #cbd5e1', borderRadius: '4px', textAlign: 'center', fontSize: '12px' }} /></div>
              <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#475569' }}>Lề Trên</label><input type="number" min="0" max="30" value={printConfig.marginTop || 5} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, marginTop: Number(e.target.value) } }))} style={{ width: '35px', border: '1px solid #cbd5e1', borderRadius: '4px', textAlign: 'center', fontSize: '12px' }} /></div>
              <div><label style={{ fontSize: '11px', fontWeight: 'bold', display: 'block', color: '#475569' }}>Lề Dưới</label><input type="number" min="0" max="30" value={printConfig.marginBottom || 5} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, marginBottom: Number(e.target.value) } }))} style={{ width: '35px', border: '1px solid #cbd5e1', borderRadius: '4px', textAlign: 'center', fontSize: '12px' }} /></div>
            </div>
          </div>

          <div style={{ marginLeft: 'auto', display: 'flex', gap: '8px', alignItems: 'center' }}>
            <div style={{ position: 'relative' }}>
              <button onClick={() => setShowExportMenu(!showExportMenu)} style={{ padding: '8px 12px', backgroundColor: '#8b5cf6', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><CheckSquare size={16} /> Chọn Lớp Xuất/In ({(activeState.selectedExportClasses || []).length})</button>
              {showExportMenu && (
                <div style={{ position: 'absolute', top: '100%', right: 0, marginTop: '8px', backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '15px', width: '220px', boxShadow: '0 10px 25px rgba(0,0,0,0.2)', zIndex: 100 }}>
                  <div style={{ display: 'flex', gap: '8px', marginBottom: '10px' }}>
                    <button onClick={() => handleToggleAllExportClasses(true)} style={{ flex: 1, padding: '4px', fontSize: '11px', backgroundColor: '#e0f2fe', color: '#0369a1', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>Chọn tất cả</button>
                    <button onClick={() => handleToggleAllExportClasses(false)} style={{ flex: 1, padding: '4px', fontSize: '11px', backgroundColor: '#fee2e2', color: '#991b1b', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>Bỏ chọn</button>
                  </div>
                  <div style={{ maxHeight: '200px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    {(activeState.dividedClasses || []).map((cls, idx) => (
                      <label key={idx} style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', cursor: 'pointer', color: '#334155' }}><input type="checkbox" checked={(activeState.selectedExportClasses || []).includes(idx)} onChange={() => handleToggleExportClass(idx)} />{cls.className}</label>
                    ))}
                  </div>
                </div>
              )}
            </div>

            <button onClick={() => setShowStatsModal(true)} style={{ padding: '8px 12px', backgroundColor: '#ec4899', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><BarChart2 size={16} /> Xem Thống Kê</button>
            <button onClick={() => window.print()} style={{ padding: '8px 12px', backgroundColor: '#0ea5e9', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><Printer size={16} /> In PDF ({classesToExport.length} Lớp)</button>
            <button onClick={exportToExcel} style={{ padding: '8px 12px', backgroundColor: '#16a34a', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><Download size={16} /> Xuất Excel</button>
          </div>
        </div>

        {/* THIẾT LẬP TIÊU ĐỀ IN */}
        <div className="no-print" style={{ display: 'flex', flexWrap: 'wrap', gap: '12px', padding: '10px 20px', backgroundColor: '#f8fafc', borderBottom: '1px solid #e2e8f0' }}>
            <div style={{ display: 'flex', alignItems: 'center', padding: '4px 8px', backgroundColor: '#e2e8f0', borderRadius: '6px', gap: '10px' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', fontWeight: 'bold', cursor: 'pointer', color: '#0f172a' }}><input type="checkbox" checked={printConfig.showHeader !== false} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, showHeader: e.target.checked } }))} />{printConfig.showHeader !== false ? <Eye size={16} color="#0284c7" /> : <EyeOff size={16} color="#64748b" />} Hiện Hành chính</label>
              <div style={{width: '1px', height: '14px', backgroundColor: '#cbd5e1'}}></div>
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', fontWeight: 'bold', cursor: 'pointer', color: '#0f172a' }}><input type="checkbox" checked={printConfig.showFooter !== false} onChange={e => updateActiveGradeState(p => ({ printConfig: { ...p.printConfig, showFooter: e.target.checked } }))} />{printConfig.showFooter !== false ? <Eye size={16} color="#0284c7" /> : <EyeOff size={16} color="#64748b" />} Hiện Chữ ký</label>
            </div>
            <button onClick={clearGlobalConfig} style={{ marginLeft: 'auto', padding: '6px 12px', backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px' }} title="Bỏ cấu hình đã lưu và quay về mặc định ban đầu"><RotateCcw size={14} /> Bỏ Cấu Hình Đã Lưu</button>
        </div>

        <div className="no-print" style={{ display: 'flex', flexWrap: 'wrap', gap: '12px', padding: '15px 20px', backgroundColor: '#e2e8f0' }}>
           <div style={{ flex: '1 1 200px', backgroundColor: '#fff', padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e1' }}>
              <label style={{ fontWeight: 'bold', fontSize: '12px', color: '#0284c7', display: 'flex', alignItems: 'center', gap: '4px' }}><Type size={14} /> Tiêu đề chính văn bản:</label>
              <input value={activeState.docTitle || ''} onChange={e => updateActiveGradeState({ docTitle: e.target.value })} placeholder="DANH SÁCH HỌC SINH" style={{ width: '100%', padding: '5px', border: '1px solid #0284c7', borderRadius: '4px', marginTop: '4px', fontWeight: 'bold', color: '#0284c7' }} />
           </div>
           <div style={{ flex: '1 1 200px', backgroundColor: '#fff', padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e1' }}>
              <label style={{ fontWeight: 'bold', fontSize: '12px', color: '#0284c7', display: 'flex', alignItems: 'center', gap: '4px' }}><Type size={14} /> Năm học / Tiêu đề phụ:</label>
              <input value={activeState.academicYear || ''} onChange={e => updateActiveGradeState({ academicYear: e.target.value })} placeholder="NĂM HỌC 2024 - 2025" style={{ width: '100%', padding: '5px', border: '1px solid #0284c7', borderRadius: '4px', marginTop: '4px', fontWeight: 'bold', color: '#0284c7' }} />
           </div>

           {printConfig.showHeader !== false && (
             <>
               <div style={{ flex: '1 1 220px' }}><label style={{ fontWeight: 'bold', fontSize: '12px', color: '#334155' }}>Cơ quan chủ quản:</label><input value={docFormat.coQuanChuQuan || ''} onChange={e => updateActiveGradeState(p => ({ docFormat: { ...p.docFormat, coQuanChuQuan: e.target.value } }))} style={{ width: '100%', padding: '5px', border: '1px solid #94a3b8', borderRadius: '4px', marginTop: '4px' }} /></div>
               <div style={{ flex: '1 1 180px' }}><label style={{ fontWeight: 'bold', fontSize: '12px', color: '#334155' }}>Cơ quan ban hành:</label><input value={docFormat.coQuanBanHanh || ''} onChange={e => updateActiveGradeState(p => ({ docFormat: { ...p.docFormat, coQuanBanHanh: e.target.value } }))} style={{ width: '100%', padding: '5px', border: '1px solid #94a3b8', borderRadius: '4px', marginTop: '4px', fontWeight: 'bold' }} /></div>
               <div style={{ flex: '1 1 220px' }}><label style={{ fontWeight: 'bold', fontSize: '12px', color: '#334155' }}>Địa danh, ngày tháng:</label><input value={docFormat.diaDanh || ''} onChange={e => updateActiveGradeState(p => ({ docFormat: { ...p.docFormat, diaDanh: e.target.value } }))} style={{ width: '100%', padding: '5px', border: '1px solid #94a3b8', borderRadius: '4px', marginTop: '4px', fontStyle: 'italic' }} /></div>
             </>
           )}

           {printConfig.showFooter !== false && (
             <>
              <div style={{ flex: '1 1 140px' }}><label style={{ fontWeight: 'bold', fontSize: '12px', color: '#334155' }}>Chức danh ký:</label><input value={docFormat.chucVu || ''} onChange={e => updateActiveGradeState(p => ({ docFormat: { ...p.docFormat, chucVu: e.target.value } }))} style={{ width: '100%', padding: '5px', border: '1px solid #94a3b8', borderRadius: '4px', marginTop: '4px', fontWeight: 'bold' }} /></div>
              <div style={{ flex: '1 1 180px' }}><label style={{ fontWeight: 'bold', fontSize: '12px', color: '#334155' }}>Nội dung dưới chữ ký:</label><input value={docFormat.nguoiKy || ''} onChange={e => updateActiveGradeState(p => ({ docFormat: { ...p.docFormat, nguoiKy: e.target.value } }))} style={{ width: '100%', padding: '5px', border: '1px solid #94a3b8', borderRadius: '4px', marginTop: '4px' }} /></div>
             </>
           )}
        </div>

        {renderModals()}

        <div className="print-content-container" style={{ marginTop: '20px' }}>
          {classesToExport.length > 0 ? (
            classesToExport.map((cls, idx) => {
              const rowLimit = printConfig.fixedRowCount || 54;
              const visibleStudents = (cls.students || []).filter(s => !(activeState.hiddenPrintIds || []).includes(getStudentId(s)));
              const totalStudents = visibleStudents.length;
              const chunks = [];
              if (totalStudents === 0) { chunks.push([]); } 
              else { for (let i = 0; i < totalStudents; i += rowLimit) { chunks.push(visibleStudents.slice(i, i + rowLimit)); } }
              const isLastClass = idx === classesToExport.length - 1;
              return (
                <div key={idx} className="class-print-wrapper">
                  {chunks.map((chunk, chunkIdx) => {
                    const isLastPageOfLastClass = isLastClass && (chunkIdx === chunks.length - 1);
                    return renderA4Page({...cls, students: visibleStudents}, chunk, chunkIdx + 1, chunks.length, isLastPageOfLastClass);
                  })}
                </div>
              );
            })
          ) : (
            <div style={{ backgroundColor: '#fff', padding: '50px', textAlign: 'center', color: '#64748b', borderRadius: '10px', margin: '0 20px' }}>
              <Printer size={48} color="#94a3b8" style={{ margin: '0 auto 10px auto', display: 'block' }} />
              <h3>Chưa chọn Lớp nào để In</h3>
              <p>Vui lòng bấm nút <strong>"Chọn Lớp Xuất/In"</strong> trên thanh công cụ màu tím để chọn danh sách hiển thị.</p>
            </div>
          )}
        </div>
      </div>
    );
  }

  // =======================================================================
  // XỬ LÝ GIAO DIỆN DÀNH CHO KHÁCH VÃNG LAI (HỌC SINH / PHỤ HUYNH)
  // =======================================================================
  if (isCheckingAuth) return <div style={{textAlign: 'center', marginTop: '100px', fontSize: '20px', color: '#5a3d2b', fontWeight: 'bold'}}>Đang tải dữ liệu từ trường...</div>;
  
  if (!session) {
    if (showAdminLogin) {
      return (
        <div style={{ position: 'relative', minHeight: '100vh', backgroundColor: '#fcf8f2' }}>
          <button onClick={() => setShowAdminLogin(false)} style={{ position: 'absolute', top: 20, left: 20, padding: '10px 15px', backgroundColor: '#fff', color: '#5a3d2b', border: '2px solid #f0e6d2', borderRadius: '12px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', zIndex: 10 }}>
            ← Quay lại trang Tra Cứu
          </button>
          <Login onLocalAdminLogin={() => setIsLocalAdmin(true)} />
        </div>
      );
    }

    return (
      <div style={{ minHeight: '100vh', backgroundColor: '#fcf8f2', display: 'flex', flexDirection: 'column', alignItems: 'center', paddingTop: '20px', paddingBottom: '50px', fontFamily: '"Nunito", "Segoe UI", "Comic Sans MS", "Times New Roman", sans-serif', paddingLeft: '15px', paddingRight: '15px', boxSizing: 'border-box', position: 'relative', overflow: 'hidden' }}>
        
        {/* Nền trang trí dễ thương (Floating icons) */}
        <div style={{ position: 'absolute', top: '5%', left: '8%', fontSize: '40px', opacity: 0.6, zIndex: 0 }}>⭐</div>
        <div style={{ position: 'absolute', top: '15%', right: '10%', fontSize: '50px', opacity: 0.5, zIndex: 0 }}>📚</div>
        <div style={{ position: 'absolute', bottom: '15%', left: '10%', fontSize: '60px', opacity: 0.4, zIndex: 0 }}>📘</div>
        <div style={{ position: 'absolute', top: '45%', right: '5%', fontSize: '35px', opacity: 0.6, zIndex: 0 }}>🌟</div>
        <div style={{ position: 'absolute', bottom: '30%', right: '15%', fontSize: '45px', opacity: 0.5, zIndex: 0 }}>📐</div>

        <style>{`
          .public-box { padding: 35px; position: relative; z-index: 2; }
          .cute-input:focus { border-color: #38bdf8 !important; box-shadow: 0 0 0 4px rgba(56, 189, 248, 0.2) !important; outline: none; }
          .btn-hover:hover { background-color: #3e9cb1 !important; transform: scale(1.02); }
          @media (max-width: 600px) {
            .public-box { padding: 25px !important; }
            .school-title { font-size: 22px !important; }
            .agency-title { font-size: 13px !important; }
            .main-title { font-size: 24px !important; }
          }
        `}</style>
        
        <div style={{ width: '100%', maxWidth: '550px', display: 'flex', justifyContent: 'flex-end', marginBottom: '15px', position: 'relative', zIndex: 2 }}>
          <button onClick={() => setShowAdminLogin(true)} style={{ padding: '10px 16px', backgroundColor: '#4b88d2', color: '#fff', border: 'none', borderRadius: '12px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px', boxShadow: '0 4px 8px rgba(75, 136, 210, 0.3)', transition: 'transform 0.2s' }}>
            <UserCheck size={16} /> Giáo viên Đăng nhập
          </button>
        </div>

        <div style={{ textAlign: 'center', marginBottom: '25px', padding: '0 10px', position: 'relative', zIndex: 2 }}>
          <div className="agency-title" style={{ fontSize: '14px', fontWeight: 'bold', color: '#6b4c3a', textTransform: 'uppercase', marginBottom: '6px', letterSpacing: '0.5px' }}>UỶ BAN NHÂN DÂN XÃ XUÂN THỚI SƠN</div>
          <div className="school-title" style={{ fontSize: '26px', fontWeight: '900', color: '#5a3d2b', textTransform: 'uppercase', marginBottom: '15px', letterSpacing: '0.5px' }}>TRƯỜNG THCS TÂN XUÂN</div>
          
          {/* Cụm biểu tượng sách vở */}
          <div style={{ display: 'flex', justifyContent: 'center', gap: '12px', fontSize: '48px', margin: '15px 0' }}>
             📚 🌍 ✏️
          </div>
          
          <h1 className="main-title" style={{ color: '#5a3d2b', margin: '10px 0 8px 0', fontSize: '28px', fontWeight: '900' }}>Cổng Thông Tin Tìm Lớp Của Em</h1>
          <p style={{ color: '#6b4c3a', fontSize: '17px', margin: 0, fontWeight: 'bold' }}>{activeState.academicYear || "Năm học 2026 - 2027"}</p>
        </div>

        <div className="public-box" style={{ background: '#fffcf6', border: '2px solid #fbe6d4', borderRadius: '28px', boxShadow: '0 15px 35px rgba(90, 61, 43, 0.08)', width: '100%', maxWidth: '550px', boxSizing: 'border-box' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '2px dashed #fcd34d', paddingBottom: '15px', marginBottom: '25px' }}>
             <h2 style={{ margin: 0, color: '#5a3d2b', fontSize: '22px', display: 'flex', alignItems: 'center', gap: '10px', fontWeight: '900' }}>
               <Search size={26} color="#5a3d2b" strokeWidth={2.5} /> Tìm Lớp Của Em
             </h2>
             <div style={{ fontSize: '32px' }}>👦🏻👧🏻</div>
          </div>
          
          <form onSubmit={handlePublicSearch} style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            <div>
              <label style={{ display: 'block', marginBottom: '10px', fontSize: '16px', fontWeight: 'bold', color: '#5a3d2b' }}>1. Họ và tên của em (*)</label>
              <input type="text" className="cute-input" placeholder="VD: Nguyễn Văn A..." value={publicSearchName} onChange={e => setPublicSearchName(e.target.value)} required style={{ width: '100%', padding: '16px', borderRadius: '14px', border: '2px solid #7dd3fc', fontSize: '16px', boxSizing: 'border-box', backgroundColor: '#fff', color: '#0f172a', transition: 'all 0.2s' }} />
            </div>
            
            <div>
              <label style={{ display: 'block', marginBottom: '10px', fontSize: '16px', fontWeight: 'bold', color: '#5a3d2b' }}>2. Ngày sinh của em (Không bắt buộc)</label>
              <input type="text" className="cute-input" placeholder="VD: 15/08/2012 hoặc 2012" value={publicSearchDob} onChange={e => setPublicSearchDob(e.target.value)} style={{ width: '100%', padding: '16px', borderRadius: '14px', border: '2px solid #7dd3fc', fontSize: '16px', boxSizing: 'border-box', backgroundColor: '#fff', color: '#0f172a', transition: 'all 0.2s' }} />
              <div style={{ fontSize: '13px', color: '#94a3b8', marginTop: '10px', fontStyle: 'italic' }}>⭐ Nhập thêm ngày sinh nếu tên của em bị trùng với bạn khác.</div>
            </div>

            <button type="submit" className="btn-hover" style={{ width: '100%', padding: '18px', backgroundColor: '#4cb5c5', color: '#fff', border: 'none', borderRadius: '22px', fontSize: '18px', fontWeight: '900', cursor: 'pointer', marginTop: '10px', boxShadow: '0 6px 15px rgba(76, 181, 197, 0.3)', transition: 'all 0.2s', display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '8px' }}>
              <Search size={22} strokeWidth={3} /> TÌM LỚP NGAY
            </button>
          </form>

          {publicSearchResult && (
            <div style={{ marginTop: '30px', paddingTop: '25px', borderTop: '2px dashed #fcd34d' }}>
              {publicSearchResult.error ? (
                <div style={{ padding: '18px', backgroundColor: '#fef2f2', color: '#b91c1c', borderRadius: '14px', fontSize: '15px', textAlign: 'center', border: '2px solid #fca5a5', fontWeight: 'bold' }}>
                  {publicSearchResult.error}
                </div>
              ) : (
                <div>
                  <div style={{ fontSize: '16px', color: '#166534', fontWeight: 'bold', marginBottom: '18px', textAlign: 'center' }}>🎉 Chúc mừng! Em đã tìm thấy:</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '15px', maxHeight: '350px', overflowY: 'auto' }}>
                    {publicSearchResult.data.map((hs, idx) => (
                      <div key={idx} style={{ padding: '22px', backgroundColor: '#f0fdf4', border: '2px solid #bbf7d0', borderRadius: '18px', position: 'relative' }}>
                        <div style={{ fontSize: '20px', fontWeight: '900', color: '#14532d', marginBottom: '10px' }}>{hs.name}</div>
                        <div style={{ fontSize: '15px', color: '#166534', display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid #bbf7d0', paddingBottom: '12px', marginBottom: '12px' }}>
                          <span>Sinh ngày: <strong>{hs.dob}</strong></span>
                          <span>Giới tính: <strong>{hs.gender}</strong></span>
                        </div>
                        <div style={{ fontSize: '16px', color: '#166534', textAlign: 'center', marginTop: '12px', fontWeight: 'bold' }}>
                          Em đã được phân vào:
                        </div>
                        <div style={{ marginTop: '8px', fontSize: '32px', fontWeight: '900', color: '#4cb5c5', textAlign: 'center', textShadow: '1px 1px 0 rgba(0,0,0,0.05)' }}>
                          LỚP {hs.className}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
        
        <div style={{ marginTop: '40px', fontSize: '13px', color: '#a89f91', textAlign: 'center', padding: '0 20px', position: 'relative', zIndex: 2, fontWeight: 'bold' }}>
          © 2026 Bản quyền thuộc về Trường THCS Tân Xuân<br/>
          Phát triển và bảo mật bởi Tác giả: Ngô Tấn Đạt
        </div>
      </div>
    );
  }

  // RETURN GIAO DIỆN CHÍNH (VIEW: MAIN)
  return (
    <div className="app-main-container" style={{ fontFamily: '"Times New Roman", Times, serif', backgroundColor: '#f1f5f9', minHeight: '100vh' }}>
      <style>{`
        .app-main-container { padding: 20px; }
        .top-toolbar { display: flex; justify-content: space-between; align-items: center; background-color: #fff; padding: 10px 20px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); margin-bottom: 20px; flex-wrap: wrap; gap: 10px; }
        .top-actions, .top-nav { display: flex; gap: 8px; flex-wrap: wrap; }
        .grade-tabs-container { display: flex; gap: 10px; justify-content: center; margin-bottom: 20px; }
        .main-content-grid { display: grid; grid-template-columns: 320px 1fr; gap: 20px; align-items: start; }
        * { font-family: "Times New Roman", Times, serif; }
        /* ĐIỀU CHỈNH GIAO DIỆN CHO ĐIỆN THOẠI */
        @media (max-width: 768px) {
          .app-main-container { padding: 10px; }
          .top-toolbar { flex-direction: column; padding: 15px; align-items: stretch; }
          .top-actions, .top-nav { flex-direction: column; width: 100%; }
          .top-actions button, .top-nav button, .top-actions input, .top-nav input { width: 100%; justify-content: center; margin-bottom: 5px; }
          .grade-tabs-container { flex-wrap: wrap; gap: 8px; }
          .grade-tabs-container button { flex: 1 1 45%; padding: 10px 5px !important; justify-content: center; flex-direction: column; text-align: center; }
          .grade-tabs-container button div { margin-top: 4px; font-size: 14px; }
          .main-content-grid { grid-template-columns: 1fr; gap: 15px; }
          .mobile-action-group { flex-direction: column; gap: 8px !important; align-items: stretch !important; }
          .mobile-action-group button, .mobile-action-group select { width: 100%; justify-content: center; margin: 0 !important; }
        }
      `}</style>
      
      <div className="top-toolbar">
        <div className="top-actions">
          <button onClick={exportSessionFile} style={{ padding: '8px 12px', backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><FileDown size={16} /> Xuất File phiên (.json)</button>
          <button onClick={() => jsonInputRef.current.click()} style={{ padding: '8px 12px', backgroundColor: '#64748b', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><FileUp size={16} /> Mở File phiên (.json)</button>
          <input type="file" ref={jsonInputRef} accept=".json" onChange={importSessionFile} style={{ display: 'none' }} />
          <button onClick={() => reImportExcelRef.current.click()} style={{ padding: '8px 12px', backgroundColor: '#8b5cf6', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}><Upload size={16} /> Import ngược Excel (Khối {activeGrade})</button>
          <input type="file" ref={reImportExcelRef} accept=".xlsx, .xls" onChange={handleImportExportedExcel} style={{ display: 'none' }} />
        </div>

        <div className="top-nav">
          <button onClick={() => setShowLogModal(true)} style={{ padding: '9px 18px', backgroundColor: '#3b82f6', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}><History size={18} /> Nhật ký thay đổi</button>
          <button onClick={undoLastAction} disabled={!activeState.actionHistory || activeState.actionHistory.length === 0} style={{ padding: '9px 18px', backgroundColor: (!activeState.actionHistory || activeState.actionHistory.length === 0) ? '#cbd5e1' : '#6366f1', color: '#fff', border: 'none', borderRadius: '6px', cursor: (!activeState.actionHistory || activeState.actionHistory.length === 0) ? 'not-allowed' : 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', transition: 'all 0.2s' }}><Undo size={18} /> Hoàn tác</button>
          <button style={{ padding: '9px 18px', backgroundColor: '#0284c7', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}><Home size={18} /> Quản lý & Phân lớp</button>
          <button onClick={() => setCurrentView('PRINT')} style={{ padding: '9px 18px', backgroundColor: '#f59e0b', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}><FileText size={18} /> Chỉnh File In / Xuất PDF</button>
        </div>
      </div>

      <h1 style={{ textAlign: 'center', color: '#0f172a', margin: '0 0 15px 0', fontSize: 'clamp(20px, 4vw, 32px)' }}>🏫 Hệ Thống Xếp Lớp Trường THCS Tân Xuân - Tác giả Ngô Tấn Đạt</h1>
      
      {/* THÔNG TIN TÀI KHOẢN */}
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '15px', marginBottom: '20px', flexWrap: 'wrap' }}>
          <span style={{ fontSize: '13px', fontWeight: 'bold', color: userRole === 'admin' ? '#16a34a' : '#ea580c' }}>
            Xin chào: {session?.user?.email} | Quyền: {userRole === 'admin' ? 'QUẢN TRỊ (ADMIN)' : 'CHỈ XEM (VIEWER)'}
          </span>
          <button 
            onClick={() => (isLocalAdmin ? setIsLocalAdmin(false) : supabase.auth.signOut())} 
            style={{ background: '#ef4444', color: 'white', padding: '4px 10px', border: 'none', borderRadius: '5px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
            Đăng xuất
          </button>
      </div>
      
      <div className="grade-tabs-container" style={{ alignItems: 'center', flexWrap: 'wrap' }}>
        {/* --- BẮT ĐẦU: KHỐI THỐNG KÊ TỔNG TOÀN TRƯỜNG --- */}
        <div style={{ backgroundColor: '#ecfdf5', border: '2px solid #10b981', padding: '10px 15px', borderRadius: '10px', display: 'flex', alignItems: 'center', gap: '8px', boxShadow: '0 2px 5px rgba(16, 185, 129, 0.2)', marginRight: '10px', transition: 'all 0.3s ease' }}>
          <span style={{ fontSize: '24px' }}>🏫</span>
          <div style={{ textAlign: 'left' }}>
            <div style={{ fontSize: '11px', color: '#065f46', fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Tổng Toàn Trường
            </div>
            <div style={{ fontSize: '18px', color: '#047857', fontWeight: '900', lineHeight: '1.2' }}>
              {totalStudentsWholeSchool ? totalStudentsWholeSchool.toLocaleString('vi-VN') : 0} HS
            </div>
          </div>
        </div>
        {/* --- KẾT THÚC: KHỐI THỐNG KÊ TỔNG TOÀN TRƯỜNG --- */}

        {['6', '7', '8', '9'].map(gKey => {
          const gData = grades[gKey] || {};
          let studentCount = 0;
          if (gData.dividedClasses && gData.dividedClasses.length > 0) {
            const assignedCount = gData.dividedClasses.reduce((total, cls) => total + (cls.students || []).length, 0);
            const assignedIds = new Set();
            gData.dividedClasses.forEach(c => (c.students || []).forEach(s => assignedIds.add(getStudentId(s))));
            const unassignedCount = (gData.data || []).filter(s => !assignedIds.has(getStudentId(s))).length;
            studentCount = assignedCount + unassignedCount;
          } else {
            studentCount = (gData.data || []).length;
          }
          const classCount = gData.dividedClasses ? gData.dividedClasses.length : 0;
          const isActive = activeGrade === gKey;

          return (
            <button key={gKey} onClick={() => setActiveGrade(gKey)} style={{ padding: '12px 20px', backgroundColor: isActive ? '#0284c7' : '#fff', color: isActive ? '#fff' : '#334155', border: isActive ? '2px solid #0284c7' : '1px solid #cbd5e1', borderRadius: '10px', fontWeight: 'bold', fontSize: '15px', cursor: 'pointer', boxShadow: isActive ? '0 4px 12px rgba(2, 132, 199, 0.3)' : '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', alignItems: 'center', gap: '8px', transition: 'all 0.2s ease' }}>
              <Layers3 size={20} color={isActive ? '#fff' : '#0284c7'} />
              <div style={{ textAlign: 'left' }}>
                <div>KHỐI {gKey}</div>
                <div style={{ fontSize: '11px', fontWeight: 'normal', opacity: 0.9 }}>{studentCount > 0 ? `${studentCount} HS ${classCount > 0 ? `• ${classCount} Lớp` : ''}` : 'Chưa nạp DL'}</div>
              </div>
            </button>
          );
        })}
      </div>

      <div className="main-content-grid">
    {/* === BODY TRÁI === */}
    <div style={{ display: 'flex', flexDirection: 'column', gap: '15px', position: 'sticky', top: '10px' }}>
      
      <div style={{ backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '10px', padding: '15px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
        <h3 style={{ margin: '0 0 10px 0', fontSize: '15px', color: '#0369a1' }}>1. Nạp File Excel cho Khối {activeGrade}</h3>
        <div style={{ display: 'flex', gap: '8px' }}>
          <label style={{ flex: 1, backgroundColor: '#0284c7', color: 'white', padding: '10px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '5px' }}>
            <PlusCircle size={16} /> Nạp Excel Khối {activeGrade}
            <input type="file" multiple accept=".xlsx, .xls" onChange={handleFileUpload} style={{ display: 'none' }} />
          </label>
          {(activeState.data || []).length > 0 && (
            <button 
              onClick={() => { 
                if(window.confirm(`Thầy có chắc chắn muốn xóa toàn bộ dữ liệu Excel của Khối ${activeGrade} không?`)) {
                  addLog(`[Khối ${activeGrade}] 🗑️ Xóa trắng toàn bộ dữ liệu Excel của khối`); 
                  updateActiveGradeState(createInitialGradeState(activeGrade)); 
                }
              }} 
              style={{ backgroundColor: '#ef4444', color: 'white', border: 'none', padding: '10px', borderRadius: '6px', cursor: 'pointer' }} 
              title="Xóa dữ liệu khối này"
            >
              <Trash2 size={16} />
            </button>
          )}
        </div>

        {(activeState.fileList || []).length > 0 && (
          <div style={{ marginTop: '10px', fontSize: '12px', color: '#334155' }}>
            <div><strong>Đã nạp Khối {activeGrade}:</strong> {(activeState.data || []).length} HS</div>
            <div style={{ display: 'flex', gap: '6px', marginTop: '6px' }}>
              <button onClick={() => setShowPreview(!showPreview)} style={{ flex: 1, backgroundColor: '#e0f2fe', color: '#0369a1', border: 'none', padding: '4px 8px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                <Eye size={12} style={{ verticalAlign: 'middle', marginRight: '4px' }} /> {showPreview ? "Ẩn Bảng Excel" : "Xem Bảng Excel"}
              </button>
              <button onClick={() => setUnassignedModalOpen(true)} style={{ flex: 1, backgroundColor: '#fef08a', color: '#b45309', border: 'none', padding: '4px 8px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                <UserPlus size={12} style={{ verticalAlign: 'middle', marginRight: '4px' }} /> HS Chưa Xếp
              </button>
            </div>
          </div>
        )}
      </div>

      {(activeState.data || []).length > 0 && (
        <div style={{ backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '10px', padding: '15px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
            <h3 style={{ margin: 0, fontSize: '15px', color: '#0369a1' }}>2. Ô Thứ Tự Ưu Tiên (Khối {activeGrade})</h3>
            <button onClick={addPriority} disabled={(activeState.priorityList || []).length >= (activeState.excelColumns || []).length} style={{ backgroundColor: '#10b981', color: 'white', border: 'none', padding: '4px 8px', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>+ Thêm Ô</button>
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '250px', overflowY: 'auto' }}>
            {(activeState.priorityList || []).map((selectedCol, index) => (
              <div key={index} style={{ display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: index < 2 ? '#f0f9ff' : '#f8fafc', padding: '6px 8px', borderRadius: '6px', border: index < 2 ? '1px solid #0284c7' : '1px solid #e2e8f0' }}>
                <span style={{ fontSize: '11px', fontWeight: 'bold', width: '95px', color: index < 2 ? '#0284c7' : '#64748b' }}>UT {index + 1} {index < 2 ? '(Cân bằng)' : '(Dịch vụ)'}:</span>
                <select value={selectedCol} onChange={(e) => updatePriority(index, e.target.value)} style={{ padding: '4px', borderRadius: '4px', border: '1px solid #94a3b8', flex: 1, fontSize: '12px', fontWeight: 'bold' }}>
                  {(activeState.excelColumns || []).map((col) => (<option key={col} value={col} disabled={(activeState.priorityList || []).includes(col) && col !== selectedCol}>{col}</option>))}
                </select>
                <button onClick={() => movePriority(index, -1)} disabled={index === 0} style={{ padding: '2px 4px', cursor: 'pointer' }}><ArrowUp size={12} /></button>
                <button onClick={() => movePriority(index, 1)} disabled={index === (activeState.priorityList || []).length - 1} style={{ padding: '2px 4px', cursor: 'pointer' }}><ArrowDown size={12} /></button>
                <button onClick={() => removePriority(index)} style={{ backgroundColor: '#fee2e2', color: '#ef4444', border: 'none', padding: '2px 6px', borderRadius: '4px', cursor: 'pointer' }}><Trash2 size={12} /></button>
              </div>
            ))}
          </div>
        </div>
      )}

      {(activeState.data || []).length > 0 && (
        <div style={{ backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '10px', padding: '15px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
          <h3 style={{ margin: '0 0 10px 0', fontSize: '15px', color: '#0369a1' }}>3. Khởi Tạo Phân Lớp (Khối {activeGrade})</h3>
          <div style={{ marginBottom: '10px', fontSize: '12px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <strong>Sĩ số chuẩn mong muốn:</strong>
            <input type="number" min="10" max="60" value={activeState.maxStudentsPerClass || 54} onChange={(e) => updateActiveGradeState({ maxStudentsPerClass: parseInt(e.target.value) || 54 })} style={{ padding: '4px', width: '50px', fontSize: '13px', fontWeight: 'bold', textAlign: 'center', border: '2px solid #0ea5e9', borderRadius: '4px' }} />
          </div>

          {unassignedStudents.length === 0 ? (
            <div style={{ backgroundColor: '#dcfce7', color: '#166534', padding: '10px', borderRadius: '6px', fontSize: '13px', textAlign: 'center', fontWeight: 'bold' }}>Tất cả học sinh đã được xếp lớp!</div>
          ) : (
            <>
              <div style={{ fontSize: '12px', color: '#ef4444', fontWeight: 'bold', marginBottom: '8px' }}>Có {unassignedStudents.length} HS chưa được xếp lớp:</div>
              <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: '6px', marginBottom: '10px' }}>
                {Object.keys(unassignedByLopCu).map((lopCu, idx) => {
                  const count = unassignedByLopCu[lopCu].length;
                  return (
                    <div key={idx} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 8px', borderBottom: '1px solid #f1f5f9', backgroundColor: selectedMixGroups.includes(lopCu) ? '#e0f2fe' : '#fff' }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', cursor: 'pointer', flex: 1, fontWeight: 'bold', color: '#334155' }}>
                        <input type="checkbox" checked={selectedMixGroups.includes(lopCu)} onChange={(e) => { if (e.target.checked) setSelectedMixGroups([...selectedMixGroups, lopCu]); else setSelectedMixGroups(selectedMixGroups.filter(l => l !== lopCu)); }} />
                        {lopCu === 'NONE' ? "Chưa có Lớp Cũ" : `Lớp ${lopCu}`} ({count} HS)
                      </label>
                      {!selectedMixGroups.includes(lopCu) && (
                        <button onClick={() => handleKeepIntactGroup(lopCu)} style={{ backgroundColor: '#10b981', color: '#fff', border: 'none', padding: '4px 8px', borderRadius: '4px', fontSize: '10px', fontWeight: 'bold', cursor: 'pointer' }} title="Bê toàn bộ học sinh nhóm này thành 1 lớp mới độc lập">Bê nguyên lớp</button>
                      )}
                    </div>
                  )
                })}
              </div>
              <button onClick={handleMixSelectedGroups} disabled={selectedMixGroups.length === 0} style={{ width: '100%', padding: '10px', backgroundColor: selectedMixGroups.length > 0 ? '#0284c7' : '#94a3b8', color: '#fff', border: 'none', borderRadius: '6px', cursor: selectedMixGroups.length > 0 ? 'pointer' : 'not-allowed', fontWeight: 'bold', fontSize: '13px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                <RefreshCw size={16} /> GOM DỊCH VỤ CÁC NHÓM CHỌN
              </button>
            </>
          )}
        </div>
      )}

      {(activeState.data || []).length > 0 && (
        <div style={{ backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '10px', padding: '15px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
          <h3 style={{ margin: '0 0 10px 0', fontSize: '15px', color: '#0369a1' }}>4. Kiểm Tra Dữ Liệu Khối {activeGrade}</h3>
          <button onClick={exportOriginalDataToExcel} style={{ width: '100%', padding: '10px', backgroundColor: '#8b5cf6', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', marginBottom: '10px' }}>
            <Download size={16} /> Xuất Dữ liệu Gốc Kèm Lớp Mới
          </button>
          <button onClick={() => updateActiveGradeState({ showDuplicates: true })} style={{ width: '100%', padding: '10px', backgroundColor: '#f59e0b', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', marginBottom: '10px' }}>
            <Search size={16} /> Kiểm tra trùng Họ Tên TOÀN KHỐI
          </button>
          <button onClick={() => setShowGhiChuModal(true)} style={{ width: '100%', padding: '10px', backgroundColor: '#ec4899', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
            <FileText size={16} /> Thống Kê Học Sinh Có Ghi Chú
          </button>

          {activeState.showDuplicates && (
            <div style={{ marginTop: '12px', fontSize: '12px', padding: '10px', borderRadius: '6px', backgroundColor: Object.keys(duplicateGroups).length === 0 ? '#dcfce7' : '#fee2e2', color: Object.keys(duplicateGroups).length === 0 ? '#166534' : '#991b1b', border: `1px solid ${Object.keys(duplicateGroups).length === 0 ? '#bbf7d0' : '#fecaca'}`, maxHeight: '250px', overflowY: 'auto' }}>
              {Object.keys(duplicateGroups).length === 0 ? (
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}><strong>✅ Không có học sinh trùng họ tên trong Khối {activeGrade}.</strong></div>
              ) : (
                <div>
                  <strong>⚠️ Phát hiện {Object.keys(duplicateGroups).length} nhóm tên bị trùng:</strong>
                  <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '8px', fontSize: '11px', backgroundColor: '#fff', color: '#0f172a', borderRadius: '4px', overflow: 'hidden' }}>
                    <thead>
                      <tr style={{ backgroundColor: '#fca5a5', color: '#7f1d1d', textAlign: 'left' }}>
                        <th style={{ padding: '6px', border: '1px solid #f87171' }}>Tên HS Trùng</th>
                        <th style={{ padding: '6px', border: '1px solid #f87171', width: '110px' }}>Ngày Sinh</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.keys(duplicateGroups).map((name, gIdx) => {
                        const rowColor = duplicateColorMap[name];
                        return duplicateGroups[name].map((st, sIdx) => (
                          <tr key={`${gIdx}_${sIdx}`} style={{ backgroundColor: sIdx % 2 === 0 ? '#fff' : '#f8fafc' }}>
                            <td style={{ padding: '5px 6px', border: `1px solid ${rowColor}50`, fontWeight: 'bold', color: rowColor }}>{st.HoTen || st['Họ và tên']}</td>
                            <td style={{ padding: '5px 6px', border: `1px solid ${rowColor}50`, color: rowColor }}>{getDob(st)}</td>
                          </tr>
                        ));
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>

    {/* === BODY PHẢI: KẾT QUẢ CỦA KHỐI ĐANG CHỌN === */}
    <div style={{ minWidth: 0 }}> 
      {showPreview && (activeState.data || []).length > 0 && (
        <div style={{ backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '10px', padding: '15px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <h4 style={{ margin: 0, color: '#0f172a' }}>Dữ liệu Excel gốc Khối {activeGrade} (Tổng: {(activeState.data || []).length} HS):</h4>
            <span style={{ fontSize: '12px', color: '#0284c7', fontStyle: 'italic' }}>*Xóa dấu "x" hoặc sửa chữ trực tiếp tại đây, các bảng khác tự động cập nhật!</span>
          </div>
          <div style={{ overflowX: 'auto', maxHeight: '500px', overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: '6px' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
              <thead style={{ position: 'sticky', top: 0, zIndex: 2 }}>
                <tr style={{ backgroundColor: '#0284c7', color: '#fff' }}>
                  <th style={{ border: '1px solid #cbd5e1', padding: '6px', textAlign: 'center', width: '40px', whiteSpace: 'nowrap', backgroundColor: '#0284c7' }}>STT</th>
                  {(activeState.excelColumns || []).map(col => {
                    let minW = '80px';
                    if (col.toLowerCase().includes('tên') || col.toLowerCase().includes('hoten')) minW = '180px';
                    else if (['hl', 'hk', 'gioitinh', 'giới tính', 'maylanh', 'giaoducstem', 'avbn', 'kynangsong', 'ic3'].includes(col.toLowerCase())) minW = '50px';
                    return <th key={col} style={{ border: '1px solid #cbd5e1', padding: '6px', textAlign: 'center', whiteSpace: 'nowrap', minWidth: minW, backgroundColor: '#0284c7' }}>{col}</th>;
                  })}
                  <th style={{ border: '1px solid #cbd5e1', padding: '6px', textAlign: 'center', whiteSpace: 'nowrap', backgroundColor: '#0284c7', width: '60px' }}>Thao tác</th>
                </tr>
              </thead>
              <tbody>
                {(activeState.data || []).map((row, idx) => (
                  <tr key={idx} style={{ backgroundColor: idx % 2 === 0 ? '#fff' : '#f8fafc' }}>
                    <td style={{ border: '1px solid #e2e8f0', padding: '4px 6px', textAlign: 'center', fontWeight: 'bold' }}>{idx + 1}</td>
                    {(activeState.excelColumns || []).map(col => {
                      const isName = col.toLowerCase().includes('tên') || col.toLowerCase().includes('hoten');
                      const isSmall = ['hl', 'hk', 'gioitinh', 'giới tính', 'maylanh', 'giaoducstem', 'avbn', 'kynangsong', 'ic3'].includes(col.toLowerCase());
                      return (
                        <td key={col} style={{ border: '1px solid #e2e8f0', padding: '2px', whiteSpace: 'nowrap', textAlign: isSmall ? 'center' : 'left' }}>
                          <input 
                            type="text" 
                            value={row[col] !== undefined ? row[col] : ''} 
                            onChange={(e) => handleCellEdit(idx, col, e.target.value)}
                            style={{ width: '100%', minWidth: isName ? '170px' : (isSmall ? '40px' : '70px'), border: '1px solid transparent', background: 'transparent', outline: 'none', padding: '4px', fontSize: '12px', textAlign: isSmall ? 'center' : 'left', borderRadius: '4px', transition: 'border 0.2s' }}
                            onFocus={e => e.target.style.border = '1px solid #93c5fd'}
                            onBlur={e => e.target.style.border = '1px solid transparent'}
                          />
                        </td>
                      );
                    })}
                    <td style={{ border: '1px solid #e2e8f0', padding: '4px 6px', textAlign: 'center' }}>
                      <button onClick={() => handleDeleteOriginalStudent(idx)} style={{ backgroundColor: '#ef4444', color: '#fff', border: 'none', borderRadius: '4px', padding: '4px 8px', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold' }}>Xóa</button>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot style={{ position: 'sticky', bottom: 0, backgroundColor: '#e0f2fe', zIndex: 2, boxShadow: '0 -2px 5px rgba(0,0,0,0.05)' }}>
                <tr>
                  <td style={{ border: '1px solid #bae6fd', padding: '6px', textAlign: 'center', fontWeight: 'bold', color: '#0369a1' }}>Mới</td>
                  {(activeState.excelColumns || []).map(col => (
                    <td key={`new_${col}`} style={{ border: '1px solid #bae6fd', padding: '4px' }}>
                      <input type="text" value={newStudentData[col] || ''} onChange={e => setNewStudentData({ ...newStudentData, [col]: e.target.value })} placeholder={`Nhập ${col}...`} style={{ width: '100%', boxSizing: 'border-box', padding: '4px', fontSize: '11px', border: '1px solid #93c5fd', borderRadius: '4px' }} />
                    </td>
                  ))}
                  <td style={{ border: '1px solid #bae6fd', padding: '6px', textAlign: 'center' }}>
                    <button onClick={handleAddStudentToOriginal} style={{ backgroundColor: '#10b981', color: '#fff', border: 'none', borderRadius: '4px', padding: '4px 8px', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold', whiteSpace: 'nowrap' }}>+ Thêm HS</button>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {activeState.dividedClasses && activeState.dividedClasses.length > 0 && (
        <div style={{ backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '10px', padding: '15px', marginBottom: '15px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
          <h3 style={{ margin: '0 0 10px 0', fontSize: '15px', color: '#0369a1', display: 'flex', alignItems: 'center', gap: '6px' }}><Search size={18} /> Tra cứu nhanh Lớp của Học sinh</h3>
          <input type="text" placeholder="Nhập tên học sinh cần tìm..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} style={{ width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #94a3b8', fontSize: '14px', boxSizing: 'border-box' }} />
          {searchQuery.trim() !== "" && (
            <div style={{ marginTop: '10px', backgroundColor: '#f0f9ff', padding: '10px', borderRadius: '6px', border: '1px solid #bae6fd', maxHeight: '150px', overflowY: 'auto' }}>
              {searchResults.length > 0 ? (
                <ul style={{ margin: 0, paddingLeft: '20px', fontSize: '13px' }}>
                  {searchResults.map((res, idx) => (
                    <li key={idx} style={{ marginBottom: '6px' }}><strong>{res.name}</strong> <span style={{ color: '#0284c7', fontSize: '12px' }}>({res.dob})</span> - Giới tính: {res.gender} - Đang ở: <span style={{ color: '#ef4444', fontWeight: 'bold', fontSize: '14px' }}>{res.className}</span></li>
                  ))}
                </ul>
              ) : (<span style={{ fontSize: '13px', color: '#64748b' }}>Không tìm thấy học sinh nào khớp với từ khóa "{searchQuery}".</span>)}
            </div>
          )}
        </div>
      )}

      {activeState.dividedClasses ? (
        <div>
          <div style={{ backgroundColor: '#fff', padding: '15px', borderRadius: '10px', border: '1px solid #cbd5e1', marginBottom: '15px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
            <div>
              <h2 style={{ margin: '0 0 4px 0', fontSize: '18px', color: '#0f172a' }}>Danh Sách Các Lớp Khối {activeGrade} ({activeState.dividedClasses.length} lớp)</h2>
              <span style={{ fontSize: '13px', color: '#64748b' }}>Tích vào ô <strong>[ ] Lớp</strong> để xóa dư hoặc chia lại</span>
            </div>
            <div className="mobile-action-group" style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <button onClick={checkDuplicateInClasses} style={{ padding: '8px 12px', backgroundColor: '#e11d48', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', boxShadow: '0 2px 4px rgba(225, 29, 72, 0.2)' }}><UserCheck2 size={16} /> Kiểm trùng Họ Tên</button>
              <button onClick={exportToExcel} style={{ padding: '9px 12px', backgroundColor: '#16a34a', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}><Download size={16} /> Excel</button>
            </div>
          </div>
          
          <div className="mobile-action-group" style={{ backgroundColor: '#fff', padding: '10px 15px', borderRadius: '10px', border: '1px solid #cbd5e1', marginBottom: '15px', display: 'flex', alignItems: 'center', gap: '10px' }}>
              <select value={activeState.reDivideOption || 'ALL'} onChange={(e) => updateActiveGradeState({ reDivideOption: e.target.value })} disabled={(activeState.selectedClassIndices || []).length < 1} style={{ padding: '8px 10px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '13px', fontWeight: 'bold', outline: 'none', opacity: (activeState.selectedClassIndices || []).length > 0 ? 1 : 0.5 }}>
                <option value="ALL">Gom Lấp đầy: TỔNG DỊCH VỤ</option>
                {(activeState.priorityList || []).slice(2).map(col => (<option key={col} value={col}>Gom Lấp đầy: {col}</option>))}
              </select>
              <button onClick={processReDivisionSelected} disabled={(activeState.selectedClassIndices || []).length < 1} style={{ padding: '9px 12px', backgroundColor: (activeState.selectedClassIndices || []).length > 0 ? '#f59e0b' : '#cbd5e1', color: '#fff', border: 'none', borderRadius: '6px', cursor: (activeState.selectedClassIndices || []).length > 0 ? 'pointer' : 'not-allowed', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}><RefreshCw size={16} /> Chia lại Lớp đang chọn</button>
              <button onClick={deleteSelectedClasses} disabled={(activeState.selectedClassIndices || []).length < 1} style={{ padding: '9px 12px', backgroundColor: (activeState.selectedClassIndices || []).length > 0 ? '#ef4444' : '#cbd5e1', color: '#fff', border: 'none', borderRadius: '6px', cursor: (activeState.selectedClassIndices || []).length > 0 ? 'pointer' : 'not-allowed', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}><Trash2 size={16} /> Xóa Lớp đang chọn</button>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '15px' }}>
            {activeState.dividedClasses.map((cls, classIdx) => {
              const nam = (cls.students || []).filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() === 'nam').length;
              const nu = (cls.students || []).filter(s => (s.GioiTinh || '').toString().trim().toLowerCase() !== 'nam').length;
              const maxCap = activeState.maxStudentsPerClass || 54;
              const isFull = (cls.students || []).length === maxCap;
              const isOver = (cls.students || []).length > maxCap;
              const isLastClass = classIdx === activeState.dividedClasses.length - 1;
              const isSelected = (activeState.selectedClassIndices || []).includes(classIdx); 

              const hlT = countAttr(cls.students, 'HL', 'T'); const hlK = countAttr(cls.students, 'HL', 'K'); const hlD = countAttr(cls.students, 'HL', 'Đ'); const hlCD = countAttr(cls.students, 'HL', 'CĐ');
              const hkT = countAttr(cls.students, 'HK', 'T'); const hkK = countAttr(cls.students, 'HK', 'K'); const hkD = countAttr(cls.students, 'HK', 'Đ');

              return (
                <div key={classIdx} style={{ backgroundColor: '#fff', border: isSelected ? '2px solid #f59e0b' : (isOver ? '2px solid #ef4444' : (isFull ? '2px solid #10b981' : (isLastClass ? '2px solid #0ea5e9' : '1px solid #cbd5e1'))), borderRadius: '10px', padding: '12px', boxShadow: isSelected ? '0 0 0 3px rgba(245, 158, 11, 0.2)' : '0 2px 4px rgba(0,0,0,0.05)' }}>
                  
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #e2e8f0', paddingBottom: '6px', marginBottom: '8px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <span onClick={() => toggleSelectClass(classIdx)} style={{ cursor: 'pointer' }}>{isSelected ? <CheckSquare color="#f59e0b" size={20} /> : <Square color="#94a3b8" size={20} />}</span>
                      {editingClassName.classIdx === classIdx ? (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                          <input type="text" value={editingClassName.newName} onChange={e => setEditingClassName({...editingClassName, newName: e.target.value})} style={{ width: '80px', padding: '2px 4px', fontWeight: 'bold', fontSize: '14px', color: '#0369a1', border: '1px solid #0ea5e9', borderRadius: '4px' }} autoFocus />
                          <button onClick={saveClassName} style={{ background: '#10b981', color: '#fff', border: 'none', borderRadius: '4px', padding: '4px 6px', cursor: 'pointer', fontWeight: 'bold' }}>Lưu</button>
                        </div>
                      ) : (
                        <h3 style={{ margin: 0, color: '#0369a1', fontSize: '15px', display: 'flex', alignItems: 'center', gap: '6px' }}>{cls.className} <button onClick={() => setEditingClassName({ classIdx, newName: cls.className })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', padding: 0 }} title="Sửa tên lớp"><Edit2 size={14} /></button></h3>
                      )}
                    </div>

                    <div>
                      <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#0f172a', marginRight: '6px' }}>{(cls.students || []).length}/{maxCap} HS</span>
                      {isFull && <span style={{ fontSize: '10px', backgroundColor: '#10b981', color: '#fff', padding: '2px 5px', borderRadius: '4px', fontWeight: 'bold' }}>FULL 100%</span>}
                      {isOver && <span style={{ fontSize: '10px', backgroundColor: '#ef4444', color: '#fff', padding: '2px 5px', borderRadius: '4px', fontWeight: 'bold' }}>VƯỢT SĨ SỐ</span>}
                    </div>
                  </div>

                  <div style={{ backgroundColor: '#f8fafc', padding: '10px', borderRadius: '6px', fontSize: '12px', marginBottom: '10px', border: '1px solid #e2e8f0' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 'bold', color: '#0f172a', marginBottom: '6px' }}><span>👨 Nam: {nam} &nbsp;|&nbsp; 👩 Nữ: {nu}</span></div>
                    <div style={{ marginBottom: '4px', color: '#334155' }}><strong>HL:</strong> {formatStatText(`T(${hlT}) - K(${hlK}) - Đ(${hlD}) - CĐ(${hlCD})`, 'HL')}</div>
                    <div style={{ marginBottom: '6px', color: '#334155' }}><strong>HK:</strong> {formatStatText(`T(${hkT}) - K(${hkK}) - Đ(${hkD})`, 'HK')}</div>
                    <div style={{ marginTop: '6px', borderTop: '1px dashed #cbd5e1', paddingTop: '6px', display: 'flex', gap: '8px', flexWrap: 'nowrap', overflowX: 'auto', whiteSpace: 'nowrap', color: '#0369a1', fontWeight: 'bold', fontSize: '11px' }}>
                      <span style={{ color: '#0284c7' }}>❄️ Máy lạnh: {(cls.students || []).filter(hasMayLanh).length}</span>
                      <span>🧪 STEM: {(cls.students || []).filter(hasStem).length}</span>
                      <span>🗣️ AVBN: {(cls.students || []).filter(hasAVBN).length}</span>
                      <span>🧘 KNS: {(cls.students || []).filter(hasKNS).length}</span>
                      <span>💻 IC3: {(cls.students || []).filter(hasIC3).length}</span>
                    </div>
                  </div>

                  <div style={{ maxHeight: '250px', overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: '6px' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', tableLayout: 'auto' }}>
                      <tbody>
                        {(cls.students || []).map((st, studentIdx) => {
                          const stNameClean = cleanName(st.HoTen || st['Họ và tên']);
                          const dupColor = duplicateColorMap[stNameClean]; 
                          
                          return (
                            <tr key={studentIdx} style={{ backgroundColor: studentIdx % 2 === 0 ? '#fff' : '#f8fafc', borderBottom: '1px solid #f1f5f9' }}>
                              <td style={{ padding: '6px', whiteSpace: 'nowrap', textAlign: 'center', width: '30px' }}>
                                {studentIdx + 1}. <button onClick={() => openStudentDetailModal(classIdx, studentIdx, st)} style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '0 0 0 4px', color: '#0284c7' }} title="Xem chi tiết, Chuyển lớp, Xóa"><Edit2 size={12} /></button>
                              </td>
                              
                              <td style={{ padding: '6px', fontWeight: 'bold', color: dupColor || 'inherit', whiteSpace: 'nowrap' }}>
                                {st.HoTen || st['Họ và tên']} 
                                <span style={{ color: '#64748b', marginLeft: '6px', marginRight: '6px', fontSize: '11px', fontWeight: 'normal' }}>({getDob(st)})</span>
                                {hasMayLanh(st) && <span title="Máy lạnh" style={{marginRight: '2px', fontSize: '11px'}}>❄️</span>}
                                {hasStem(st) && <span title="Giáo dục STEM" style={{marginRight: '2px', fontSize: '11px'}}>🧪</span>}
                                {hasAVBN(st) && <span title="Anh văn bản ngữ" style={{marginRight: '2px', fontSize: '11px'}}>🗣️</span>}
                                {hasKNS(st) && <span title="Kỹ năng sống" style={{marginRight: '2px', fontSize: '11px'}}>🧘</span>}
                                {hasIC3(st) && <span title="Tin học IC3" style={{marginRight: '2px', fontSize: '11px'}}>💻</span>}
                                {getGhiChu(st) && <span style={{ marginLeft: '4px', fontSize: '11px', color: '#ef4444', fontWeight: 'bold', fontStyle: 'italic' }}>[{getGhiChu(st)}]</span>}
                              </td>
                              
                              <td style={{ padding: '6px', color: dupColor || 'inherit', whiteSpace: 'nowrap', textAlign: 'center' }}>{st.GioiTinh}</td>
                              <td style={{ padding: '6px', color: dupColor || 'inherit', whiteSpace: 'nowrap', textAlign: 'center' }}>HL: {st.HL}</td>
                              <td style={{ padding: '6px', color: dupColor || 'inherit', whiteSpace: 'nowrap', textAlign: 'center' }}>HK: {st.HK}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ) : (
        <div style={{ backgroundColor: '#fff', border: '2px dashed #cbd5e1', borderRadius: '10px', padding: '50px', textAlign: 'center', color: '#64748b' }}>
          <Layers size={48} color="#94a3b8" style={{ margin: '0 auto 10px auto', display: 'block' }} />
          <h3>Khối {activeGrade} chưa có dữ liệu phân lớp</h3>
          <p>Vui lòng nạp File Excel bên trái và bấm <strong>"GOM DỊCH VỤ LẤP ĐẦY KHỐI {activeGrade}"</strong></p>
        </div>
      )}
    </div>
  </div>

  {renderModals()}
</div>
  );
}