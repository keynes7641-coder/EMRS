/**
 * 設備借用管理系統 - Google Apps Script 後端程式碼 (Code.gs)
 * 
 * 部署說明：
 * 1. 建立一個全新的 Google 試算表 (Google Sheets)。
 * 2. 點擊選單「擴充功能」->「Apps Script」。
 * 3. 將此程式碼完整複製並取代 Code.gs 內容。
 * 4. 點擊「執行」選單中的 `setupDatabase` 函式，初始化三個 Sheet 頁籤與預設範例資料。
 * 5. 點擊右上角「部署」->「新增部署」:
 *    - 種類選擇：「網頁應用程式 (Web App)」
 *    - 說明：設備借用系統 API
 *    - 誰可以存取 (Who has access)：【所有人 (Anyone)】  <-- 必須設定為所有人！
 * 6. 複製獲得的「網頁應用程式 URL」，貼入前端系統的 API 網址設定中。
 */

// 工作表名稱定義
const SHEETS = {
  EQUIPMENT: 'Equipment',
  RECORDS: 'Records',
  ADMINS: 'Admins'
};

/**
 * 初始化資料庫：建立 3 個工作表、寫入標題列與範例資料
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Equipment 工作表
  let eqSheet = ss.getSheetByName(SHEETS.EQUIPMENT);
  if (!eqSheet) {
    eqSheet = ss.insertSheet(SHEETS.EQUIPMENT);
  }
  eqSheet.clear();
  eqSheet.appendRow(['id', 'name', 'category', 'total_qty', 'available_qty', 'location', 'status']);
  eqSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#E0F2FE');
  
  // 寫入預設設備
  const defaultEquipment = [
    ['EQ-1001', 'MacBook Pro 16吋 M2', '筆記型電腦', 5, 4, '創客中心 A101', '可借用'],
    ['EQ-1002', 'Sony Alpha 7 IV 相機', '攝影器材', 3, 2, '視聽教室 B202', '可借用'],
    ['EQ-1003', 'Epson 無線投影機 EH-TW740', '視聽設備', 4, 4, '器材室 C301', '可借用'],
    ['EQ-1004', 'DJI Ronin RS3 穩定器', '攝影器材', 2, 1, '視聽教室 B202', '可借用'],
    ['EQ-1005', 'iPad Pro 12.9吋 (含Apple Pencil)', '平板電腦', 8, 6, '創客中心 A101', '可借用'],
    ['EQ-1006', 'Rode Wireless GO II 無線麥克風', '音響設備', 4, 3, '器材室 C301', '可借用']
  ];
  defaultEquipment.forEach(row => eqSheet.appendRow(row));

  // 2. Records 工作表 (加入 equipment_name 欄位永久保存設備名稱)
  let recSheet = ss.getSheetByName(SHEETS.RECORDS);
  if (!recSheet) {
    recSheet = ss.insertSheet(SHEETS.RECORDS);
  }
  recSheet.clear();
  recSheet.appendRow(['record_id', 'user_name', 'user_email', 'equipment_id', 'equipment_name', 'borrow_qty', 'borrow_time', 'expected_return_time', 'actual_return_time', 'status']);
  recSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#FEF3C7');

  // 寫入範例借用紀錄
  const now = new Date();
  const dateStr = formatDate(now);
  const futureStr = formatDate(new Date(now.getTime() + 7 * 86400000));
  
  const defaultRecords = [
    ['REC-20260801-01', '王小明', 'ming@example.com', 'EQ-1001', 'MacBook Pro 16吋 M2', 1, dateStr, futureStr, '', '已出借'],
    ['REC-20260801-02', '陳大文', 'david@example.com', 'EQ-1002', 'Sony Alpha 7 IV 相機', 1, dateStr, futureStr, '', '待審核']
  ];
  defaultRecords.forEach(row => recSheet.appendRow(row));

  // 3. Admins 工作表
  let admSheet = ss.getSheetByName(SHEETS.ADMINS);
  if (!admSheet) {
    admSheet = ss.insertSheet(SHEETS.ADMINS);
  }
  admSheet.clear();
  admSheet.appendRow(['email', 'role', 'password']);
  admSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#DCFCE7');

  // 寫入預設管理者 (預設密碼 admin123)
  admSheet.appendRow(['admin@example.com', 'SuperAdmin', 'admin123']);
  admSheet.appendRow(['manager@example.com', 'Manager', 'admin123']);
  admSheet.appendRow(['wzmit@wzm.kh.edu.tw', 'SuperAdmin', 'admin123']);

  // 刪除預設的 Sheet1 (若存在且非唯一)
  const sheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('工作表1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  Logger.log('資料庫初始化成功！');
  return 'Database setup completed successfully!';
}

// -------------------------------------------------------------
// RESTful API 進入點: GET 與 POST
// -------------------------------------------------------------

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getEquipment';

    if (action === 'setup') {
      const msg = setupDatabase();
      return jsonResponse({ success: true, message: msg });
    }

    if (action === 'getEquipment') {
      const data = getSheetObjects(SHEETS.EQUIPMENT);
      return jsonResponse({ success: true, data: data });
    }

    if (action === 'getRecords') {
      const userEmail = e.parameter ? e.parameter.email : null;
      let records = getSheetObjects(SHEETS.RECORDS);
      if (userEmail) {
        records = records.filter(r => String(r.user_email).toLowerCase() === userEmail.toLowerCase());
      }
      
      // 建立設備 Map 作為舊相容回退
      const equipmentMap = {};
      getSheetObjects(SHEETS.EQUIPMENT).forEach(eq => {
        equipmentMap[eq.id] = eq;
      });

      records = records.map(rec => {
        // 優先讀取紀錄中保存的 equipment_name；若舊欄位無設備名稱，才從 Map 查找，最後回退到 equipment_id
        let eqName = rec.equipment_name;
        if (!eqName || eqName === '未知設備') {
          eqName = equipmentMap[rec.equipment_id] ? equipmentMap[rec.equipment_id].name : (rec.equipment_id || '已報廢刪除設備');
        }

        return {
          ...rec,
          equipment_name: eqName,
          location: equipmentMap[rec.equipment_id] ? equipmentMap[rec.equipment_id].location : ''
        };
      });

      return jsonResponse({ success: true, data: records });
    }

    if (action === 'getAdmins') {
      const admins = getSheetObjects(SHEETS.ADMINS).map(a => ({ email: a.email, role: a.role }));
      return jsonResponse({ success: true, data: admins });
    }

    if (action === 'cleanUpOldRecords') {
      const result = cleanUpOldRecords();
      return jsonResponse(result);
    }

    if (action === 'checkBorrowReminders') {
      const result = checkBorrowReminders();
      return jsonResponse(result);
    }

    return jsonResponse({ success: false, message: '無效的 action 指令' });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return jsonResponse({ success: false, message: '系統忙碌中，請稍後重試' });
  }

  try {
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    const action = (e && e.parameter && e.parameter.action) || body.action;

    // 1. 管理員驗證登入
    if (action === 'verifyAdmin') {
      const { email, password } = body;
      const admins = getSheetObjects(SHEETS.ADMINS);
      const matched = admins.find(a => String(a.email).toLowerCase() === String(email).toLowerCase() && String(a.password) === String(password));
      
      if (matched) {
        return jsonResponse({ success: true, data: { email: matched.email, role: matched.role } });
      } else {
        return jsonResponse({ success: false, message: '帳號或密碼錯誤' });
      }
    }

    // 2. 提出借用申請 (將設備名稱 equipment_name 永久寫入 Records 工作表)
    if (action === 'createBorrow') {
      const { user_name, user_email, equipment_id, equipment_name, borrow_qty, expected_return_time } = body;
      if (!user_name || !user_email || !equipment_id || !borrow_qty) {
        return jsonResponse({ success: false, message: '請填寫完整借用資訊' });
      }

      const eqList = getSheetObjects(SHEETS.EQUIPMENT);
      const eq = eqList.find(item => String(item.id) === String(equipment_id));
      
      // 確定借用當下的設備名稱 (優先使用前端傳入或歷史查得之名稱)
      const finalEqName = equipment_name || (eq ? eq.name : '未知設備');

      // 檢查申請人是否有逾期未歸還之設備
      const userRecords = getSheetObjects(SHEETS.RECORDS).filter(r => String(r.user_email).toLowerCase() === String(user_email).toLowerCase());
      const todayMidnight = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()).getTime();

      const overdueRecords = userRecords.filter(r => {
        if (r.status !== '已出借' || !r.expected_return_time) return false;
        const expDate = new Date(r.expected_return_time.replace(/-/g, '/'));
        if (isNaN(expDate.getTime())) return false;
        const expMidnight = new Date(expDate.getFullYear(), expDate.getMonth(), expDate.getDate()).getTime();
        return expMidnight < todayMidnight;
      });

      if (overdueRecords.length > 0) {
        return jsonResponse({
          success: false,
          message: `借用失敗！您目前尚有 ${overdueRecords.length} 筆設備到期未歸還，借用權限已暫時凍結，請先辦理歸還手續。`
        });
      }

      if (!eq) {
        return jsonResponse({ success: false, message: '找不到對應設備' });
      }

      const qty = parseInt(borrow_qty, 10);
      if (isNaN(qty) || qty <= 0) {
        return jsonResponse({ success: false, message: '借用數量必須大於 0' });
      }

      if (eq.available_qty < qty) {
        return jsonResponse({ success: false, message: `可借用數量不足 (現有: ${eq.available_qty})` });
      }

      const recordId = 'REC-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 1000);
      const nowStr = formatDate(new Date());

      const recSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RECORDS);
      
      // 寫入 10 個欄位：record_id, user_name, user_email, equipment_id, equipment_name, borrow_qty, borrow_time, expected_return_time, actual_return_time, status
      recSheet.appendRow([
        recordId,
        user_name,
        user_email,
        equipment_id,
        finalEqName,
        qty,
        nowStr,
        expected_return_time || '',
        '',
        '待審核'
      ]);

      return jsonResponse({ success: true, message: '借用申請已提交，等待管理者審核', record_id: recordId });
    }

    // 3. 核准借出
    if (action === 'approveBorrow') {
      const { record_id } = body;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const recSheet = ss.getSheetByName(SHEETS.RECORDS);
      const eqSheet = ss.getSheetByName(SHEETS.EQUIPMENT);

      const records = getSheetObjects(SHEETS.RECORDS);
      const recIndex = records.findIndex(r => String(r.record_id) === String(record_id));
      if (recIndex === -1) {
        return jsonResponse({ success: false, message: '找不到借用紀錄' });
      }

      const rec = records[recIndex];
      if (rec.status !== '待審核') {
        return jsonResponse({ success: false, message: `目前狀態為「${rec.status}」，無法執行核准` });
      }

      const equipments = getSheetObjects(SHEETS.EQUIPMENT);
      const eqIndex = equipments.findIndex(e => String(e.id) === String(rec.equipment_id));
      if (eqIndex === -1) {
        return jsonResponse({ success: false, message: '找不到對應設備' });
      }

      const eq = equipments[eqIndex];
      const borrowQty = parseInt(rec.borrow_qty, 10);
      if (eq.available_qty < borrowQty) {
        return jsonResponse({ success: false, message: `即時剩餘數量不足 (剩餘: ${eq.available_qty}, 欲借: ${borrowQty})` });
      }

      // 扣減庫存
      const newAvail = eq.available_qty - borrowQty;
      eqSheet.getRange(eqIndex + 2, 5).setValue(newAvail); // Col 5: available_qty

      // 更新紀錄狀態 (檢查與動態尋找 status 欄位位置)
      const recHeaders = recSheet.getRange(1, 1, 1, recSheet.getLastColumn()).getValues()[0];
      const statusCol = recHeaders.indexOf('status') + 1;
      recSheet.getRange(recIndex + 2, statusCol > 0 ? statusCol : 10).setValue('已出借');

      return jsonResponse({ success: true, message: '已成功核准借出' });
    }

    // 4. 駁回申請
    if (action === 'rejectBorrow') {
      const { record_id } = body;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const recSheet = ss.getSheetByName(SHEETS.RECORDS);

      const records = getSheetObjects(SHEETS.RECORDS);
      const recIndex = records.findIndex(r => String(r.record_id) === String(record_id));
      if (recIndex === -1) {
        return jsonResponse({ success: false, message: '找不到借用紀錄' });
      }

      const recHeaders = recSheet.getRange(1, 1, 1, recSheet.getLastColumn()).getValues()[0];
      const statusCol = recHeaders.indexOf('status') + 1;
      recSheet.getRange(recIndex + 2, statusCol > 0 ? statusCol : 10).setValue('已駁回');

      return jsonResponse({ success: true, message: '已駁回借用申請' });
    }

    // 5. 確認歸還 (庫存自動加回)
    if (action === 'returnEquipment') {
      const { record_id } = body;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const recSheet = ss.getSheetByName(SHEETS.RECORDS);
      const eqSheet = ss.getSheetByName(SHEETS.EQUIPMENT);

      const records = getSheetObjects(SHEETS.RECORDS);
      const recIndex = records.findIndex(r => String(r.record_id) === String(record_id));
      if (recIndex === -1) {
        return jsonResponse({ success: false, message: '找不到借用紀錄' });
      }

      const rec = records[recIndex];
      if (rec.status !== '已出借') {
        return jsonResponse({ success: false, message: `目前狀態為「${rec.status}」，無需歸還` });
      }

      const equipments = getSheetObjects(SHEETS.EQUIPMENT);
      const eqIndex = equipments.findIndex(e => String(e.id) === String(rec.equipment_id));

      if (eqIndex !== -1) {
        const eq = equipments[eqIndex];
        const borrowQty = parseInt(rec.borrow_qty, 10);
        const newAvail = Math.min(eq.total_qty, eq.available_qty + borrowQty);
        eqSheet.getRange(eqIndex + 2, 5).setValue(newAvail);
      }

      const nowStr = formatDate(new Date());
      const recHeaders = recSheet.getRange(1, 1, 1, recSheet.getLastColumn()).getValues()[0];
      const actualReturnCol = recHeaders.indexOf('actual_return_time') + 1;
      const statusCol = recHeaders.indexOf('status') + 1;

      recSheet.getRange(recIndex + 2, actualReturnCol > 0 ? actualReturnCol : 9).setValue(nowStr);
      recSheet.getRange(recIndex + 2, statusCol > 0 ? statusCol : 10).setValue('已歸還');

      return jsonResponse({ success: true, message: '已確認歸還，庫存已自動加回！' });
    }

    // 6. 新增設備
    if (action === 'addEquipment') {
      const { name, category, total_qty, available_qty, location, status } = body;
      if (!name || !category) {
        return jsonResponse({ success: false, message: '名稱與分類為必填欄位' });
      }

      const total = parseInt(total_qty, 10) || 1;
      const avail = available_qty !== undefined ? parseInt(available_qty, 10) : total;
      const eqId = 'EQ-' + Math.floor(1000 + Math.random() * 9000);

      const eqSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPMENT);
      eqSheet.appendRow([
        eqId,
        name,
        category,
        total,
        avail,
        location || '未指定',
        status || '可借用'
      ]);

      return jsonResponse({ success: true, message: '新設備已新增成功', id: eqId });
    }

    // 7. 更新設備
    if (action === 'updateEquipment') {
      const { id, name, category, total_qty, available_qty, location, status } = body;
      const eqSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPMENT);
      const equipments = getSheetObjects(SHEETS.EQUIPMENT);
      const eqIndex = equipments.findIndex(e => String(e.id) === String(id));

      if (eqIndex === -1) {
        return jsonResponse({ success: false, message: '找不到要更新的設備' });
      }

      const row = eqIndex + 2;
      if (name !== undefined) eqSheet.getRange(row, 2).setValue(name);
      if (category !== undefined) eqSheet.getRange(row, 3).setValue(category);
      if (total_qty !== undefined) eqSheet.getRange(row, 4).setValue(parseInt(total_qty, 10));
      if (available_qty !== undefined) eqSheet.getRange(row, 5).setValue(parseInt(available_qty, 10));
      if (location !== undefined) eqSheet.getRange(row, 6).setValue(location);
      if (status !== undefined) eqSheet.getRange(row, 7).setValue(status);

      return jsonResponse({ success: true, message: '設備資料已更新' });
    }

    // 8. 刪除設備 (具備未歸還出借檢查)
    if (action === 'deleteEquipment') {
      const { id } = body;
      const eqSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPMENT);
      const equipments = getSheetObjects(SHEETS.EQUIPMENT);
      const eqIndex = equipments.findIndex(e => String(e.id) === String(id));

      if (eqIndex === -1) {
        return jsonResponse({ success: false, message: '找不到要刪除的設備' });
      }

      const targetEq = equipments[eqIndex];

      if (targetEq.available_qty < targetEq.total_qty) {
        const unreturnedCount = targetEq.total_qty - targetEq.available_qty;
        return jsonResponse({
          success: false,
          message: `禁止刪除！該設備目前尚有 ${unreturnedCount} 件處於「已出借」狀態，請先辦理歸還手續。`
        });
      }

      const records = getSheetObjects(SHEETS.RECORDS);
      const activeBorrowRecords = records.filter(r => String(r.equipment_id) === String(id) && r.status === '已出借');
      if (activeBorrowRecords.length > 0) {
        return jsonResponse({
          success: false,
          message: `禁止刪除！發現該設備有 ${activeBorrowRecords.length} 筆未歸還的「已出借」紀錄。`
        });
      }

      eqSheet.deleteRow(eqIndex + 2);
      return jsonResponse({ success: true, message: '設備已成功刪除' });
    }

    // 9. 新增管理員帳號
    if (action === 'addAdmin') {
      const { email, password, role } = body;
      if (!email || !password) {
        return jsonResponse({ success: false, message: '請輸入 Email 與密碼' });
      }

      const admSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ADMINS);
      const admins = getSheetObjects(SHEETS.ADMINS);
      const exists = admins.some(a => String(a.email).toLowerCase() === String(email).toLowerCase());

      if (exists) {
        return jsonResponse({ success: false, message: '該管理員 Email 已存在' });
      }

      admSheet.appendRow([email.toLowerCase(), role || 'Manager', password]);
      return jsonResponse({ success: true, message: '已成功新增管理員帳號！' });
    }

    // 10. 修改管理員密碼 / 角色
    if (action === 'updateAdmin') {
      const { email, newPassword, role } = body;
      const admSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ADMINS);
      const admins = getSheetObjects(SHEETS.ADMINS);
      const admIndex = admins.findIndex(a => String(a.email).toLowerCase() === String(email).toLowerCase());

      if (admIndex === -1) {
        return jsonResponse({ success: false, message: '找不到要修改的管理員' });
      }

      const row = admIndex + 2;
      if (role !== undefined) admSheet.getRange(row, 2).setValue(role);
      if (newPassword !== undefined && newPassword !== '') admSheet.getRange(row, 3).setValue(newPassword);

      return jsonResponse({ success: true, message: '管理員資料與密碼已成功更新！' });
    }

    // 11. 刪除管理員帳號
    if (action === 'deleteAdmin') {
      const { email } = body;
      const admSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ADMINS);
      const admins = getSheetObjects(SHEETS.ADMINS);

      if (admins.length <= 1) {
        return jsonResponse({ success: false, message: '系統必須至少保留一個管理員帳號' });
      }

      const admIndex = admins.findIndex(a => String(a.email).toLowerCase() === String(email).toLowerCase());
      if (admIndex === -1) {
        return jsonResponse({ success: false, message: '找不到要刪除的管理員' });
      }

      admSheet.deleteRow(admIndex + 2);
      return jsonResponse({ success: true, message: '管理員帳號已成功刪除' });
    }

    // 12. 手動執行歷史紀錄清理與逾期 Email 通知
    if (action === 'cleanUpOldRecords') {
      const result = cleanUpOldRecords();
      return jsonResponse(result);
    }

    // 13. 執行設備借用到期 Email 提醒檢查 (7天前/到期日)
    if (action === 'checkBorrowReminders') {
      const result = checkBorrowReminders();
      return jsonResponse(result);
    }

    return jsonResponse({ success: false, message: '未知的 POST action 操作' });

  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------
// 工具函式
// -------------------------------------------------------------

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim());
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.join('').trim() === '') continue;

    const obj = {};
    headers.forEach((header, index) => {
      let val = row[index];
      if (val instanceof Date) {
        val = formatDate(val);
      }
      obj[header] = val;
    });
    result.push(obj);
  }

  return result;
}

function formatDate(d) {
  if (!d) return '';
  return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
}

/**
 * 歷史紀錄清理與逾期通知機制 (紀錄保留 1 年)
 * 1. 遍歷 Records 試算表中所有紀錄。
 * 2. 判斷紀錄之 borrow_time 是否超過 365 天。
 * 3. 超過 1 年者：
 *    - 若狀態為「已歸還」或「已駁回」：將該列資料刪除。
 *    - 若狀態為「待審核」或「已出借」（即未歸還）：保留紀錄，並發送 Email 通知借用者與 SuperAdmin 管理員。
 */
function cleanUpOldRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recSheet = ss.getSheetByName(SHEETS.RECORDS);
  if (!recSheet) return { success: false, message: '找不到 Records 工作表' };

  const records = getSheetObjects(SHEETS.RECORDS);
  const admins = getSheetObjects(SHEETS.ADMINS);

  // 取得 SuperAdmin 的 Email 清單
  const superAdminEmails = admins
    .filter(a => String(a.role).toLowerCase() === 'superadmin' && a.email)
    .map(a => String(a.email).trim());

  const now = new Date();
  const ONE_YEAR_MS = 365 * 24 * 60 * 60 * 1000; // 365 天 (86400000 ms * 365)

  let deletedCount = 0;
  let notifiedCount = 0;
  const notifiedDetails = [];

  // 從最底部開始向上遍歷，確保 deleteRow 不會影響尚未處理列的索引位置
  for (let i = records.length - 1; i >= 0; i--) {
    const rec = records[i];
    const rowIndex = i + 2; // 表頭標題佔據第 1 列

    if (!rec.borrow_time) continue;

    // 解析借用時間
    const borrowDate = new Date(rec.borrow_time);
    if (isNaN(borrowDate.getTime())) continue;

    const diffMs = now.getTime() - borrowDate.getTime();

    // 判斷是否超過一年 (365天)
    if (diffMs > ONE_YEAR_MS) {
      const isReturnedOrRejected = (rec.status === '已歸還' || rec.status === '已駁回');

      if (isReturnedOrRejected) {
        // 已結案之舊紀錄：直接刪除
        recSheet.deleteRow(rowIndex);
        deletedCount++;
      } else {
        // 未歸還之舊紀錄：保留紀錄並發送 Email 通知
        notifiedCount++;
        notifiedDetails.push(rec);

        // 寄送通知信
        try {
          const subject = `【EMRS 設備借用逾期通知】您有超過 1 年未歸還的設備：${rec.equipment_name || rec.equipment_id}`;
          const body = `親愛的 ${rec.user_name} 您好：\n\n` +
            `系統偵測到您於 ${rec.borrow_time} 申請借用的設備「${rec.equipment_name || rec.equipment_id}」（數量: ${rec.borrow_qty}），` +
            `至今已超過 1 年且狀態仍為「${rec.status}」。\n\n` +
            `【紀錄詳細資訊】\n` +
            `- 紀錄編號：${rec.record_id}\n` +
            `- 設備名稱：${rec.equipment_name || rec.equipment_id}\n` +
            `- 借用數量：${rec.borrow_qty}\n` +
            `- 借用時間：${rec.borrow_time}\n` +
            `- 預計歸還時間：${rec.expected_return_time || '未填寫'}\n` +
            `- 目前狀態：${rec.status}\n\n` +
            `請您儘速至器材室歸還設備或聯繫系統管理員辦理續借程序。\n` +
            `如已歸還，請通知管理者於後台更正紀錄狀態。\n\n` +
            `設備借用管理系統 (EMRS) 敬上`;

          // 收件者包含使用者以及所有 SuperAdmin
          const recipientsList = [rec.user_email, ...superAdminEmails].filter(Boolean);
          const uniqueRecipients = Array.from(new Set(recipientsList.map(e => e.toLowerCase()))).join(',');

          if (uniqueRecipients) {
            MailApp.sendEmail({
              to: uniqueRecipients,
              subject: subject,
              body: body
            });
          }
        } catch (mailErr) {
          Logger.log('發送逾期通知郵件失敗: ' + mailErr.toString());
        }
      }
    }
  }

  const resultMsg = `舊紀錄清理完成！共刪除 ${deletedCount} 筆超過1年已歸還紀錄，並向 ${notifiedCount} 筆超過1年未歸還紀錄發送了逾期 Email 通知。`;
  Logger.log(resultMsg);

  return {
    success: true,
    message: resultMsg,
    deletedCount: deletedCount,
    notifiedCount: notifiedCount,
    notifiedRecords: notifiedDetails
  };
}

/**
 * 設備借用到期提醒機制 (建議設定 GAS 每日 Time-Driven 觸發器)
 * 1. 到期日前一週 (7 天)：發送 Email 提醒借用者「距離設備借用到期只剩一週」
 * 2. 設備借用到期日 (0 天)：發送 Email 提醒借用者「今天下班前歸還設備」
 */
function checkBorrowReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recSheet = ss.getSheetByName(SHEETS.RECORDS);
  if (!recSheet) return { success: false, message: '找不到 Records 工作表' };

  const records = getSheetObjects(SHEETS.RECORDS);
  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

  let count7Days = 0;
  let countDueToday = 0;

  records.forEach(rec => {
    if (rec.status !== '已出借' || !rec.expected_return_time || !rec.user_email) return;

    // 解析預計歸還日期
    const expDate = new Date(rec.expected_return_time.replace(/-/g, '/'));
    if (isNaN(expDate.getTime())) return;

    const expMidnight = new Date(expDate.getFullYear(), expDate.getMonth(), expDate.getDate()).getTime();
    const diffDays = Math.round((expMidnight - todayMidnight) / (24 * 60 * 60 * 1000));

    if (diffDays === 7) {
      // 1. 倒數 7 天提醒
      try {
        const subject = `【EMRS 到期提醒】您借用的設備距離歸還期限只剩 1 週`;
        const body = `親愛的 ${rec.user_name} 您好：\n\n` +
          `提醒您，您申請借用的設備「${rec.equipment_name || rec.equipment_id}」（數量: ${rec.borrow_qty}）距離預計歸還日期只剩一週！\n\n` +
          `【借用詳細資訊】\n` +
          `- 紀錄編號：${rec.record_id}\n` +
          `- 設備名稱：${rec.equipment_name || rec.equipment_id}\n` +
          `- 借用數量：${rec.borrow_qty}\n` +
          `- 借用時間：${rec.borrow_time}\n` +
          `- 預計歸還日期：${rec.expected_return_time}\n\n` +
          `請記得於到期日前準備歸還，謝謝！\n\n` +
          `設備借用管理系統 (EMRS) 敬上`;

        MailApp.sendEmail({
          to: rec.user_email,
          subject: subject,
          body: body
        });
        count7Days++;
      } catch (err) {
        Logger.log('發送 7 天前提醒郵件失敗: ' + err.toString());
      }
    } else if (diffDays === 0) {
      // 2. 今日到期提醒
      try {
        const subject = `【EMRS 到期提醒】您借用的設備今日到期，請於今天下班前歸還設備`;
        const body = `親愛的 ${rec.user_name} 您好：\n\n` +
          `提醒您，您申請借用的設備「${rec.equipment_name || rec.equipment_id}」（數量: ${rec.borrow_qty}）預計歸還日期為今天（${rec.expected_return_time}）。\n\n` +
          `【借用詳細資訊】\n` +
          `- 紀錄編號：${rec.record_id}\n` +
          `- 設備名稱：${rec.equipment_name || rec.equipment_id}\n` +
          `- 借用數量：${rec.borrow_qty}\n` +
          `- 借用時間：${rec.borrow_time}\n` +
          `- 預計歸還日期：${rec.expected_return_time}\n\n` +
          `請務必於今天下班前將設備歸還至器材室，以免影響您的設備借用權限。\n\n` +
          `設備借用管理系統 (EMRS) 敬上`;

        MailApp.sendEmail({
          to: rec.user_email,
          subject: subject,
          body: body
        });
        countDueToday++;
      } catch (err) {
        Logger.log('發送今日到期提醒郵件失敗: ' + err.toString());
      }
    }
  });

  const resultMsg = `到期提醒檢查完成！共發送 7 天前提醒 ${count7Days} 封，今日到期提醒 ${countDueToday} 封。`;
  Logger.log(resultMsg);

  return {
    success: true,
    message: resultMsg,
    count7Days: count7Days,
    countDueToday: countDueToday
  };
}

