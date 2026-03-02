// Google Apps Script - 業務統計表後端 API

// 配置：設定你的 Google Sheet ID
// 從 Google Sheet 網址取得：https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
const SPREADSHEET_ID = '1_xReH9JB2XH429KlHLrtFvDjqt52Cfzpwfn7JxVOycQ';

// 設定表單回應的 sheet 名稱
const SHEET_NAME = 'Sheet1';

/**
 * 處理 GET 請求 (用于測試Web App)
 */
function doGet(e) {
    return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: '業務統計表 API 正常運作',
        timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 處理 POST 請求 (接收表單資料)
 */
function doPost(e) {
    try {
        // 解析收到的 JSON 資料
        const data = JSON.parse(e.postData.contents);
        
        // 驗證資料
        if (!validateData(data)) {
            return ContentService.createTextOutput(JSON.stringify({
                status: 'error',
                message: '資料驗證失敗'
            })).setMimeType(ContentService.MimeType.JSON);
        }
        
        // 開啟 Google Sheet
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        
        // 設定標頭列（如果第一列為空）
        if (sheet.getLastRow() === 0) {
            const headers = [
                '登入日期', '承辦人', '工作分類', '起造人', '消防核准文號', 
                '地址(地號)', '委託人', '建築物用途分類', '樓層概要', '總樓地板面積', 
                '附加簽證', '備註'
            ];
            sheet.appendRow(headers);
        }
        
        // 準備要寫入的資料列
        const row = [
            data.loginDate || '',
            data.contactPerson || '',
            data.workType || '',
            data.builder || '',
            data.firePermit || '',
            data.address || '',
            data.client || '',
            data.buildingType || '',
            data.floors || '',
            data.totalArea || '',
            data.certifications || '',
            data.notes || ''
        ];
        
        // 寫入資料
        sheet.appendRow(row);
        
        // 回傳成功訊息
        return ContentService.createTextOutput(JSON.stringify({
            status: 'success',
            message: '資料已成功儲存',
            timestamp: new Date().toISOString()
        })).setMimeType(ContentService.MimeType.JSON);
        
    } catch (error) {
        // 回傳錯誤訊息
        return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: error.toString(),
            timestamp: new Date().toISOString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * 驗證資料完整性
 */
function validateData(data) {
    // 檢查必填欄位
    const requiredFields = [
        'loginDate', 'contactPerson', 'workType', 
        'builder', 'address', 'client', 'buildingType', 
        'floors', 'totalArea'
    ];
    
    for (const field of requiredFields) {
        if (!data[field] || data[field].trim() === '') {
            return false;
        }
    }
    
    // 檢查承辦人至少選擇一個
    if (!data.contactPerson || data.contactPerson.trim() === '') {
        return false;
    }
    
    return true;
}

/**
 * 取得所有已提交的資料
 */
function getAllData() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
        const values = sheet.getDataRange().getValues();
        return values;
    } catch (error) {
        return { error: error.toString() };
    }
}
