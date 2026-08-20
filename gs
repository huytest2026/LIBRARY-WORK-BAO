const TEMPLATE_DOC_ID = '1ekR7FCEMQemjjGlGVmV9dfbEhsE29Ih5Jizy9umkod0';
const STORAGE_FOLDER_ID = '1piQfjhGZHkOtKDPtrmsIP1cq5Ksp9uJi';
const LOG_SHEET_ID = '1MQ2trQbvKsoIOsq3q8FD9mOwkH21FS-G3aM3dw3DO08';

const KHO_FILE_ID_1 = '1uVuVbNk8TlB9yx9WbJm4wRtIiB4_r6V5ym-3qlGG1Ro'; 
const KHO_FILE_ID_2 = '1KRy41q69eHLE4b0UtJw3HRdspQKJHsUcZDaN6tD1s4Q';         

function getTargetSpreadsheet() {
  try {
    if (LOG_SHEET_ID && LOG_SHEET_ID.trim() !== "") {
      return SpreadsheetApp.openById(LOG_SHEET_ID.trim());
    }
  } catch (e) {}
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return HtmlService.createTemplateFromFile('test')
      .evaluate()
      .setTitle('Hệ Thống Quản Lý Ca Trực & Sự Cố Cáp Quang')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function toCanonicalDate(dateStr) {
  if (!dateStr) return '';
  try {
    let d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
      return d.toISOString().split('T')[0];
    }
  } catch (e) {}
  return dateStr.toString().trim();
}

function parseSheetDate(dateStr) {
  if (!dateStr) return null;
  try {
    let d = new Date(dateStr);
    if (!isNaN(d.getTime())) return d;
  } catch (e) {}
  return null;
}

function getGivenName(fullName) {
  if (!fullName) return '';
  let parts = fullName.trim().split(/\s+/);
  return parts[parts.length - 1];
}

function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
    let staffList = [];
    let staffSheet = ss.getSheetByName('Danh sách nhân viên');
    if (staffSheet) {
      let lastRow = staffSheet.getLastRow();
      if (lastRow >= 2) {
        let values = staffSheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
        staffList = values.map(r => r[0]).filter(v => v && v.trim() !== '');
      }
    }
    if (staffList.length === 0) {
      staffList = ["Trần Quốc Bảo", "Nguyễn Văn A", "Lê Văn B"];
    }

    let allVatTuData = [];
    let vtSheet = ss.getSheetByName('VẬT TƯ VT');
    if (vtSheet) {
      let lastRow = vtSheet.getLastRow();
      let lastCol = vtSheet.getLastColumn();
      if (lastRow >= 2) {
        let vals = vtSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
        let headers = vals[0];
        let colMap = {};
        for (let j = 0; j < headers.length; j++) {
          colMap[headers[j].toString().trim().toUpperCase()] = j;
        }

        for (let i = 1; i < vals.length; i++) {
          let r = vals[i];
          allVatTuData.push({
            maKho: colMap['MA_KHO'] !== undefined ? (r[colMap['MA_KHO']] || '') : (r[6] || ''),
            tenKho: colMap['TEN_KHO'] !== undefined ? (r[colMap['TEN_KHO']] || '') : (r[7] || ''),
            loHang: colMap['LOHANG'] !== undefined ? (r[colMap['LOHANG']] || '') : (colMap['LO_HANG'] !== undefined ? r[colMap['LO_HANG']] : (r[10] || '')),
            maVt: colMap['MA_VT'] !== undefined ? (r[colMap['MA_VT']] || '') : (r[13] || ''),
            tenVt: colMap['TEN_VT'] !== undefined ? (r[colMap['TEN_VT']] || '') : (r[15] || ''),
            tenTat: colMap['TENTAT'] !== undefined ? (r[colMap['TENTAT']] || '') : (colMap['TEN_TAT'] !== undefined ? r[colMap['TEN_TAT']] : (r[16] || '')),
            partNumber: colMap['PARTNUMBER'] !== undefined ? (r[colMap['PARTNUMBER']] || '') : (colMap['PART_NUMBER'] !== undefined ? r[colMap['PART_NUMBER']] : ''),
            serial: colMap['SERIAL'] !== undefined ? (r[colMap['SERIAL']] || '') : (colMap['SERIAL_NO'] !== undefined ? r[colMap['SERIAL_NO']] : ''),
            linhVuc: colMap['LINH_VUC'] !== undefined ? (r[colMap['LINH_VUC']] || '') : (colMap['LINHVUC'] !== undefined ? r[colMap['LINHVUC']] : (r[8] || ''))
          });
        }
      }
    }

    let linhVucSet = new Set();
    allVatTuData.forEach(item => {
      if (item.linhVuc && item.linhVuc.trim() !== '') {
        linhVucSet.add(item.linhVuc.trim());
      }
    });
    let linhVucList = Array.from(linhVucSet).sort((a, b) => a.localeCompare(b, 'vi', { sensitivity: 'base' }));
    if (linhVucList.length === 0) {
      linhVucList = ["Vật tư Vô tuyến", "Vật tư Cáp quang", "Vật tư Truyền dẫn"];
    }

    let lastHandover = null;
    let giaoCaSheet = ss.getSheetByName('GiaoCa');
    if (giaoCaSheet) {
      let lastRow = giaoCaSheet.getLastRow();
      if (lastRow >= 2) {
        let vals = giaoCaSheet.getRange(lastRow, 1, 1, giaoCaSheet.getLastColumn()).getDisplayValues()[0];
        lastHandover = {
          ngayGiao: vals[0] || '',
          caTruc: vals[1] || '',
          nguoiGiao: vals[2] || '',
          nguoiNhan: vals[3] || '',
          congViecThuongXuyen: vals[4] || '',
          noiDungCongViec: vals[5] || '',
          noiDungTonTai: vals[6] || ''
        };
      }
    }

    return {
      success: true,
      staffList: staffList,
      linhVucList: linhVucList,
      allVatTuData: allVatTuData,
      lastHandover: lastHandover
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function processCapNhatKho(base64Data, fileName, mimeType, targetKho) {
  try {
    const decodedBytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
    
    const tempFile = DriveApp.createFile(blob);
    const tempSpreadsheet = SpreadsheetApp.open(tempFile);
    const tempSheet = tempSpreadsheet.getSheets()[0];
    const sourceData = tempSheet.getDataRange().getValues();
    
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    
    if (sourceData.length === 0) {
      return { success: false, error: 'File tải lên không có dữ liệu.' };
    }
    
    const targetSs = SpreadsheetApp.openById(LOG_SHEET_ID);
    const sheetName = "VẬT TƯ VT";
    let targetSheet = targetSs.getSheetByName(sheetName);
    
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet(sheetName);
    }
    
    targetSheet.clear();
    targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
    
    return { 
      success: true, 
      message: `Đã xóa dữ liệu cũ và cập nhật thành công ${sourceData.length} dòng vào sheet "${sheetName}"!` 
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function importFileToDataSheet(base64Data, fileName, mimeType) {
  try {
    const ss = getTargetSpreadsheet();
    if (!ss) return { success: false, error: 'Không thể kết nối với Google Sheet.' };
    
    let sheet = ss.getSheetByName('DATA');
    if (!sheet) sheet = ss.insertSheet('DATA');
    
    const decodedBytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
    
    const tempFile = DriveApp.createFile(blob);
    const tempSpreadsheet = SpreadsheetApp.open(tempFile);
    const tempSheet = tempSpreadsheet.getSheets()[0];
    const sourceData = tempSheet.getDataRange().getValues();
    
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    
    if (sourceData.length === 0) return { success: false, error: 'File tải lên không có dữ liệu.' };
    
    sheet.clear();
    sheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
    
    return { success: true, message: 'Đã nạp dữ liệu thành công vào sheet DATA!', totalRows: sourceData.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getBaoCaoTuDataFilter(criteria) {
  try {
    const ss = getTargetSpreadsheet();
    if (!ss) return { success: false, error: 'Không thể kết nối với Google Sheet.' };
    
    const sheet = ss.getSheetByName('DATA');
    if (!sheet) return { success: false, error: 'Không tìm thấy sheet "DATA". Vui lòng chọn file và bấm "Lấy số liệu" trước.' };
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { success: true, data: [], headers: [], message: 'Sheet DATA chưa có dữ liệu.' };
    
    const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = values[0]; 
    const rows = values.slice(1); 
    
    let filteredRows = [];
    let startDate = criteria.tuNgay ? new Date(criteria.tuNgay) : null;
    let endDate = criteria.denNgay ? new Date(criteria.denNgay) : null;
    if (startDate) startDate.setHours(0, 0, 0, 0);
    if (endDate) endDate.setHours(23, 59, 59, 999);
    
    rows.forEach(row => {
      let mucRow = row[2] ? row[2].toString().trim() : '';        
      let nnRow = row[8] ? row[8].toString().trim() : '';         
      let tgRow = row[10] ? row[10].toString().trim() : '';       
      
      if (criteria.muc && criteria.muc.trim() !== '' && mucRow.toLowerCase() !== criteria.muc.toLowerCase()) return;
      if (criteria.nguyenNhan && criteria.nguyenNhan.trim() !== '') {
        if (!nnRow.toLowerCase().includes(criteria.nguyenNhan.toLowerCase().trim())) return;
      }
      if (startDate && endDate && tgRow) {
        let rowDate = parseSheetDate(tgRow);
        if (rowDate && (rowDate < startDate || rowDate > endDate)) return;
      }
      filteredRows.push(row);
    });
    
    return { success: true, headers: headers, data: filteredRows };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getPhanCaByWeek(tuanBatDau) {
  try {
    const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
    let trucCaList = [];
    let tuanTraList = [];
    let staffSheet = ss.getSheetByName('Danh sách nhân viên');
    if (staffSheet) {
      const lastStaffRow = staffSheet.getLastRow();
      if (lastStaffRow >= 2) {
        const staffValues = staffSheet.getRange(2, 7, lastStaffRow - 1, 2).getDisplayValues();
        let trucCaSet = new Set();
        let tuanTraSet = new Set();
        staffValues.forEach(row => {
          if (row[0] && row[0].trim() !== '') trucCaSet.add(row[0].trim());
          if (row[1] && row[1].trim() !== '') tuanTraSet.add(row[1].trim());
        });
        trucCaList = Array.from(trucCaSet).sort((a, b) => a.localeCompare(b, 'vi', { sensitivity: 'base' }));
        tuanTraList = Array.from(tuanTraSet).sort((a, b) => a.localeCompare(b, 'vi', { sensitivity: 'base' }));
      }
    }

    let sheet = ss.getSheetByName('PhanCa');
    if (!sheet) return { success: true, tuanBatDau: tuanBatDau, data: [], trucCaList: trucCaList, tuanTraList: tuanTraList, positionNotes: {} };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, tuanBatDau: tuanBatDau, data: [], trucCaList: trucCaList, tuanTraList: tuanTraList, positionNotes: {} };
    
    const lastCol = Math.max(23, sheet.getLastColumn());
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    
    let weeksMap = {};
    values.forEach(row => {
      const rowTuan = toCanonicalDate(row[2]) || row[2].toString().trim();
      if (!rowTuan) return;
      if (!weeksMap[rowTuan]) {
        weeksMap[rowTuan] = {
          tuanBatDau: rowTuan,
          rows: [],
          positionNotes: {
            ca1: row[15] || '', ca2: row[16] || '', ca3: row[17] || '',
            ttcq1: row[18] || '', ttcq2: row[19] || '', ttcq3: row[20] || '',
            congTac: row[21] || '', nghiBu: row[22] || ''
          }
        };
      }
      weeksMap[rowTuan].rows.push({
        ngay: row[0], thu: row[1], tuanBatDau: row[2],
        ca1: row[3], ca2: row[4], ca3: row[5],
        ttcq1: row[6], ttcq2: row[7], ttcq3: row[8],
        congTac: row[9], nghiBu: row[10], ghiChu: row[11]
      });
    });

    let targetTuan = toCanonicalDate(tuanBatDau) || (tuanBatDau ? tuanBatDau.toString().trim() : '');
    let availableWeeks = Object.keys(weeksMap).sort();
    
    if (!targetTuan || !weeksMap[targetTuan]) {
      if (availableWeeks.length > 0) {
        targetTuan = availableWeeks[availableWeeks.length - 1];
      }
    }

    let weekData = [];
    let positionNotes = { ca1: '', ca2: '', ca3: '', ttcq1: '', ttcq2: '', ttcq3: '', congTac: '', nghiBu: '' };
    
    if (targetTuan && weeksMap[targetTuan]) {
      weekData = weeksMap[targetTuan].rows;
      positionNotes = weeksMap[targetTuan].positionNotes;
    }

    return { 
      success: true, 
      tuanBatDau: targetTuan, 
      data: weekData, 
      trucCaList: trucCaList, 
      tuanTraList: tuanTraList, 
      positionNotes: positionNotes 
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function savePhanCaWeek(payload) {
  try {
    const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
    let sheet = ss.getSheetByName('PhanCa');
    if (!sheet) {
      sheet = ss.insertSheet('PhanCa');
      sheet.appendRow(["Ngay", "Thu", "Tuần bắt đầu", "Ca1", "Ca2", "Ca3", "TRCQ1", "TRCQ2", "TRCQ3", "Công tác", "Nghỉ bù", "Ghi chú"]);
    }
    
    const tuanBatDau = payload.tuanBatDau;
    const daysData = payload.daysData;
    const positionNotes = payload.positionNotes || {};
    
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
      for (let i = values.length; i >= 1; i--) {
        const rowTuan = toCanonicalDate(values[i-1][2]) || values[i-1][2].toString().trim();
        const targetTuan = toCanonicalDate(tuanBatDau) || tuanBatDau.toString().trim();
        if (rowTuan === targetTuan) sheet.deleteRow(i + 1);
      }
    }
    
    daysData.forEach((d, idx) => {
      let rowValues = [
        d.ngay, d.thu, tuanBatDau, d.ca1, d.ca2, d.ca3,
        d.ttcq1, d.ttcq2, d.ttcq3, d.congTac, d.nghiBu, d.ghiChu || ''
      ];
      if (idx === 0) {
        rowValues[15] = positionNotes.ca1 || '';
        rowValues[16] = positionNotes.ca2 || '';
        rowValues[17] = positionNotes.ca3 || '';
        rowValues[18] = positionNotes.ttcq1 || '';
        rowValues[19] = positionNotes.ttcq2 || '';
        rowValues[20] = positionNotes.ttcq3 || '';
        rowValues[21] = positionNotes.congTac || '';
        rowValues[22] = positionNotes.nghiBu || '';
      }
      sheet.appendRow(rowValues);
    });
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function tinhCongNhanVien(tuNgay, denNgay) {
  try {
    const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
    const sheet = ss.getSheetByName('PhanCa');
    if (!sheet) return { success: false, error: 'Không tìm thấy dữ liệu phân ca để tính công.' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    const startCanon = toCanonicalDate(tuNgay);
    const endCanon = toCanonicalDate(denNgay);
    let stats = {};

    function addStaffRecord(name, type) {
      if (!name || name.trim() === '') return;
      const cleanName = name.trim();
      if (!stats[cleanName]) {
        stats[cleanName] = { nhanVien: cleanName, ca1: 0, ca2: 0, ca3: 0, ttcq: 0, congTac: 0, nghiBu: 0, tongCa: 0 };
      }
      if (type === 'ca1') stats[cleanName].ca1++;
      if (type === 'ca2') stats[cleanName].ca2++;
      if (type === 'ca3') stats[cleanName].ca3++;
      if (type === 'ttcq') stats[cleanName].ttcq++;
      if (type === 'congTac') stats[cleanName].congTac++;
      if (type === 'nghiBu') stats[cleanName].nghiBu++;
    }

    values.forEach(row => {
      const ngayRow = toCanonicalDate(row[0]);
      if (!ngayRow) return;
      if (startCanon && ngayRow < startCanon) return;
      if (endCanon && ngayRow > endCanon) return;

      if (row[3]) addStaffRecord(row[3], 'ca1');
      if (row[4]) addStaffRecord(row[4], 'ca2');
      if (row[5]) addStaffRecord(row[5], 'ca3');
      if (row[9]) row[9].split(',').forEach(p => addStaffRecord(p, 'congTac'));
      [6, 7, 8].forEach(colIdx => {
        if (row[colIdx]) row[colIdx].split(',').forEach(p => addStaffRecord(p, 'ttcq'));
      });
      if (row[10]) row[10].split(',').forEach(p => addStaffRecord(p, 'nghiBu'));
    });

    let resultList = Object.values(stats).map(item => {
      item.tongCa = item.ca1 + item.ca2 + item.ca3 + item.ttcq + item.congTac;
      return item;
    });

    resultList.sort((a, b) => getGivenName(a.nhanVien).localeCompare(getGivenName(b.nhanVien), 'vi'));
    return { success: true, data: resultList };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getLibraryAndSoHoaData() {
  try {
    const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
    let libraryData = [];
    let librarySheet = ss.getSheetByName('LIBRARY');
    if (librarySheet) {
      const lastRow = librarySheet.getLastRow();
      if (lastRow >= 2) {
        const values = librarySheet.getDataRange().getDisplayValues();
        const headers = values[0];
        let colMap = {};
        for (let j = 0; j < headers.length; j++) {
          colMap[headers[j].toString().trim().toUpperCase()] = j;
        }
        for (let i = 1; i < values.length; i++) {
          const row = values[i];
          libraryData.push({
            id: colMap['ID'] !== undefined ? row[colMap['ID']] : '',
            tramQuanLy: colMap['TRẠM QUẢN LÝ'] !== undefined ? row[colMap['TRẠM QUẢN LÝ']] : '',
            tuyenCap: colMap['TUYẾN CÁP'] !== undefined ? row[colMap['TUYẾN CÁP']] : '',
            link: colMap['LINK'] !== undefined ? row[colMap['LINK']] : '',
            soCapSo: colMap['SỢI CÁP SỐ'] !== undefined ? row[colMap['SỢI CÁP SỐ']] : '',
            tenTuyenQuyDinh: colMap['TÊN TUYẾN QUY ĐỊNH'] !== undefined ? row[colMap['TÊN TUYẾN QUY ĐỊNH']] : '',
            tuyenDuong: colMap['TUYẾN ĐƯỜNG'] !== undefined ? row[colMap['TUYẾN ĐƯỜNG']] : '',
            chieuDaiLink: colMap['CHIỀU DÀI LINK'] !== undefined ? row[colMap['CHIỀU DÀI LINK']] : ''
          });
        }
      }
    }

    let soHoaSheet = ss.getSheetByName('SoHoaTuyenCap');
    let soHoaData = [];
    if (soHoaSheet) {
      const lastRow = soHoaSheet.getLastRow();
      if (lastRow >= 2) {
        const values = soHoaSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < values.length; i++) {
          const r = values[i];
          soHoaData.push({
            rowIdx: i + 1,
            tramQuanLy: r[1] || '',
            tuyenQuyDinh: r[2] || '',
            link: r[3] || '',
            tuyenCap: r[4] || '',
            soCapSo: r[5] || '',
            tenTuyenQuyDinh: r[6] || '',
            tuyenDuong: r[7] || '',
            chieuDaiLink: r[8] || ''
          });
        }
      }
    }

    return {
      success: true,
      libraryData: libraryData,
      soHoaData: soHoaData
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
