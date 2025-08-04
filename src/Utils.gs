/**
 * ユーティリティ関数モジュール
 * 共通的に使用される関数を提供
 */

/**
 * 祝日判定機能
 * Master_Holidayシートを参照して指定日が祝日かどうかを判定
 * @param {Date} date - 判定対象の日付
 * @return {Object} {isHoliday: boolean, isLegalHoliday: boolean, isCompanyHoliday: boolean, name: string}
 */
function isHoliday(date) {
  return withErrorHandling(() => {
    if (!date || !(date instanceof Date)) {
      throw new Error('有効な日付オブジェクトが必要です');
    }
    
    const targetDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const holidayData = getSheetData('Master_Holiday', 'A:D');
    
    // ヘッダー行をスキップ
    for (let i = 1; i < holidayData.length; i++) {
      const row = holidayData[i];
      if (!row[0]) continue; // 空行をスキップ
      
      const holidayDate = new Date(row[0]);
      const holidayDateOnly = new Date(holidayDate.getFullYear(), holidayDate.getMonth(), holidayDate.getDate());
      
      if (targetDate.getTime() === holidayDateOnly.getTime()) {
        return {
          isHoliday: true,
          isLegalHoliday: row[2] || false,
          isCompanyHoliday: row[3] || false,
          name: row[1] || ''
        };
      }
    }
    
    return {
      isHoliday: false,
      isLegalHoliday: false,
      isCompanyHoliday: false,
      name: ''
    };
    
  }, 'Utils.isHoliday', 'LOW', {
    date: formatDate(date),
    suppressError: true,
    defaultValue: {
      isHoliday: false,
      isLegalHoliday: false,
      isCompanyHoliday: false,
      name: ''
    }
  });
}

/**
 * 時間を分単位に変換
 * @param {string|Date} timeString - 時刻文字列 (HH:MM形式) または時刻オブジェクト
 * @return {number} 分単位の時間
 */
function timeToMinutes(timeString) {
  try {
    if (!timeString) return 0;
    
    let hours, minutes;
    
    if (timeString instanceof Date) {
      hours = timeString.getHours();
      minutes = timeString.getMinutes();
    } else {
      const timeStr = timeString.toString();
      const timeParts = timeStr.split(':');
      if (timeParts.length !== 2) {
        throw new Error('時刻は HH:MM 形式で入力してください');
      }
      hours = parseInt(timeParts[0], 10);
      minutes = parseInt(timeParts[1], 10);
    }
    
    if (isNaN(hours) || isNaN(minutes) || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
      throw new Error('無効な時刻です');
    }
    
    return hours * 60 + minutes;
    
  } catch (error) {
    Logger.log('時間変換エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 分単位の時間を HH:MM 形式に変換
 * @param {number} minutes - 分単位の時間
 * @return {string} HH:MM形式の時刻文字列
 */
function minutesToTime(minutes) {
  try {
    if (typeof minutes !== 'number' || minutes < 0) {
      return '00:00';
    }
    
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;
    
    return String(hours).padStart(2, '0') + ':' + String(mins).padStart(2, '0');
    
  } catch (error) {
    Logger.log('時間フォーマットエラー: ' + error.toString());
    return '00:00';
  }
}

/**
 * 2つの時刻間の差分を分単位で計算
 * @param {string|Date} startTime - 開始時刻
 * @param {string|Date} endTime - 終了時刻
 * @return {number} 差分（分単位）
 */
function calculateTimeDifference(startTime, endTime) {
  try {
    if (!startTime || !endTime) return 0;
    
    const startMinutes = timeToMinutes(startTime);
    const endMinutes = timeToMinutes(endTime);
    
    // 日をまたぐ場合の処理
    if (endMinutes < startMinutes) {
      return (24 * 60) - startMinutes + endMinutes;
    }
    
    return endMinutes - startMinutes;
    
  } catch (error) {
    Logger.log('時間差計算エラー: ' + error.toString());
    return 0;
  }
}

/**
 * データアクセス抽象化関数
 * シートからデータを取得
 * @param {string} sheetName - シート名
 * @param {string} range - 範囲（A1記法）
 * @return {Array<Array>} シートデータの2次元配列
 */
function getSheetData(sheetName, range) {
  return withErrorHandling(() => {
    if (!sheetName) {
      throw new Error('シート名が指定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }
    
    if (!range) {
      // 範囲が指定されていない場合は、データがある範囲を自動取得
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow === 0 || lastCol === 0) {
        return []; // 空のシート
      }
      
      return sheet.getRange(1, 1, lastRow, lastCol).getValues();
    }
    
    return sheet.getRange(range).getValues();
    
  }, 'Utils.getSheetData', 'MEDIUM', {
    sheetName: sheetName,
    range: range
  });
}

/**
 * シートにデータを設定
 * @param {string} sheetName - シート名
 * @param {string} range - 範囲（A1記法）
 * @param {Array<Array>} values - 設定するデータの2次元配列
 */
function setSheetData(sheetName, range, values) {
  try {
    if (!sheetName) {
      throw new Error('シート名が指定されていません');
    }
    
    if (!values || !Array.isArray(values)) {
      throw new Error('有効なデータ配列が必要です');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }
    
    if (!range) {
      // 範囲が指定されていない場合は、A1から開始
      const numRows = values.length;
      const numCols = values[0] ? values[0].length : 0;
      
      if (numRows > 0 && numCols > 0) {
        sheet.getRange(1, 1, numRows, numCols).setValues(values);
      }
    } else {
      sheet.getRange(range).setValues(values);
    }
    
  } catch (error) {
    Logger.log(`データ設定エラー (${sheetName}, ${range}): ` + error.toString());
    throw error;
  }
}

/**
 * シートに1行のデータを追加
 * @param {string} sheetName - シート名
 * @param {Array} rowData - 追加する行データの配列
 */
function appendRowToSheet(sheetName, rowData) {
  try {
    if (!sheetName) {
      throw new Error('シート名が指定されていません');
    }
    
    if (!rowData || !Array.isArray(rowData)) {
      throw new Error('有効な行データ配列が必要です');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }
    
    sheet.appendRow(rowData);
    
  } catch (error) {
    Logger.log(`行追加エラー (${sheetName}): ` + error.toString());
    throw error;
  }
}

/**
 * 日付フォーマット関数
 * @param {Date} date - フォーマット対象の日付
 * @param {string} format - フォーマット形式 ('YYYY-MM-DD', 'YYYY/MM/DD', 'MM/DD', 'YYYY-MM')
 * @return {string} フォーマットされた日付文字列
 */
function formatDate(date, format = 'YYYY-MM-DD') {
  try {
    if (!date || !(date instanceof Date)) {
      throw new Error('有効な日付オブジェクトが必要です');
    }
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    switch (format) {
      case 'YYYY-MM-DD':
        return `${year}-${month}-${day}`;
      case 'YYYY/MM/DD':
        return `${year}/${month}/${day}`;
      case 'MM/DD':
        return `${month}/${day}`;
      case 'YYYY-MM':
        return `${year}-${month}`;
      default:
        return `${year}-${month}-${day}`;
    }
    
  } catch (error) {
    Logger.log('日付フォーマットエラー: ' + error.toString());
    return '';
  }
}

/**
 * 時刻フォーマット関数
 * @param {Date} date - フォーマット対象の日時
 * @param {string} format - フォーマット形式 ('HH:MM', 'HH:MM:SS', 'YYYY-MM-DD HH:MM')
 * @return {string} フォーマットされた時刻文字列
 */
function formatTime(date, format = 'HH:MM') {
  try {
    if (!date || !(date instanceof Date)) {
      throw new Error('有効な日付オブジェクトが必要です');
    }
    
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    
    switch (format) {
      case 'HH:MM':
        return `${hours}:${minutes}`;
      case 'HH:MM:SS':
        return `${hours}:${minutes}:${seconds}`;
      case 'YYYY-MM-DD HH:MM':
        return `${formatDate(date)} ${hours}:${minutes}`;
      default:
        return `${hours}:${minutes}`;
    }
    
  } catch (error) {
    Logger.log('時刻フォーマットエラー: ' + error.toString());
    return '';
  }
}/**
 *
 現在の日付を取得（時刻部分を除去）
 * @return {Date} 今日の日付（00:00:00）
 */
function getToday() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

/**
 * 指定した日数前/後の日付を取得
 * @param {Date} baseDate - 基準日
 * @param {number} days - 日数（負の値で過去、正の値で未来）
 * @return {Date} 計算後の日付
 */
function addDays(baseDate, days) {
  try {
    if (!baseDate || !(baseDate instanceof Date)) {
      throw new Error('有効な日付オブジェクトが必要です');
    }
    
    if (typeof days !== 'number') {
      throw new Error('日数は数値で指定してください');
    }
    
    const result = new Date(baseDate);
    result.setDate(result.getDate() + days);
    return result;
    
  } catch (error) {
    Logger.log('日付計算エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 月の最初の日を取得
 * @param {Date} date - 対象の日付
 * @return {Date} その月の1日
 */
function getFirstDayOfMonth(date) {
  try {
    if (!date || !(date instanceof Date)) {
      throw new Error('有効な日付オブジェクトが必要です');
    }
    
    return new Date(date.getFullYear(), date.getMonth(), 1);
    
  } catch (error) {
    Logger.log('月初日取得エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 月の最後の日を取得
 * @param {Date} date - 対象の日付
 * @return {Date} その月の最終日
 */
function getLastDayOfMonth(date) {
  try {
    if (!date || !(date instanceof Date)) {
      throw new Error('有効な日付オブジェクトが必要です');
    }
    
    return new Date(date.getFullYear(), date.getMonth() + 1, 0);
    
  } catch (error) {
    Logger.log('月末日取得エラー: ' + error.toString());
    throw error;
  }
}

/**
 * CSVエクスポート機能
 * @param {string} sheetName - エクスポート対象のシート名
 * @param {string} fileName - 出力ファイル名（拡張子なし）
 * @return {string} Google DriveのファイルID
 */
function exportToCsv(sheetName, fileName) {
  try {
    if (!sheetName) {
      throw new Error('シート名が指定されていません');
    }
    
    if (!fileName) {
      throw new Error('ファイル名が指定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }
    
    // シートデータを取得
    const data = getSheetData(sheetName);
    
    if (data.length === 0) {
      throw new Error('エクスポートするデータがありません');
    }
    
    // CSVデータを作成
    const csvContent = data.map(row => 
      row.map(cell => {
        // セルの値を文字列に変換し、カンマやダブルクォートをエスケープ
        const cellValue = String(cell || '');
        if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
          return '"' + cellValue.replace(/"/g, '""') + '"';
        }
        return cellValue;
      }).join(',')
    ).join('\n');
    
    // Google Driveにファイルを作成
    const blob = Utilities.newBlob(csvContent, 'text/csv', `${fileName}.csv`);
    const file = DriveApp.createFile(blob);
    
    Logger.log(`CSVファイルが作成されました: ${fileName}.csv (ID: ${file.getId()})`);
    return file.getId();
    
  } catch (error) {
    Logger.log(`CSVエクスポートエラー (${sheetName}): ` + error.toString());
    throw error;
  }
}

/**
 * 深夜勤務時間の計算
 * 22:00-05:00の時間帯の労働時間を計算
 * @param {string|Date} startTime - 開始時刻
 * @param {string|Date} endTime - 終了時刻
 * @return {number} 深夜勤務時間（分単位）
 */
function calculateNightWorkTime(startTime, endTime) {
  try {
    if (!startTime || !endTime) return 0;
    
    const startMinutes = timeToMinutes(startTime);
    const endMinutes = timeToMinutes(endTime);
    
    // 深夜時間帯の定義（22:00-05:00）
    const nightStart = 22 * 60; // 22:00 = 1320分
    const nightEnd = 5 * 60;    // 05:00 = 300分
    
    let nightWorkMinutes = 0;
    
    // 日をまたがない場合
    if (endMinutes >= startMinutes) {
      // 22:00-24:00の範囲
      if (startMinutes < nightStart && endMinutes > nightStart) {
        nightWorkMinutes += Math.min(endMinutes, 24 * 60) - nightStart;
      } else if (startMinutes >= nightStart) {
        nightWorkMinutes += endMinutes - startMinutes;
      }
      
      // 00:00-05:00の範囲
      if (startMinutes < nightEnd) {
        nightWorkMinutes += Math.min(endMinutes, nightEnd) - Math.max(startMinutes, 0);
      }
    } else {
      // 日をまたぐ場合
      // 開始日の22:00-24:00
      if (startMinutes < nightStart) {
        nightWorkMinutes += (24 * 60) - nightStart;
      } else {
        nightWorkMinutes += (24 * 60) - startMinutes;
      }
      
      // 翌日の00:00-05:00
      nightWorkMinutes += Math.min(endMinutes, nightEnd);
    }
    
    return Math.max(0, nightWorkMinutes);
    
  } catch (error) {
    Logger.log('深夜勤務時間計算エラー: ' + error.toString());
    return 0;
  }
}/**
 
* シートオブジェクトを取得
 * @param {string} sheetName - シート名
 * @return {Sheet} シートオブジェクト
 * @throws {Error} シートが見つからない場合
 */
function getSheet(sheetName) {
  try {
    if (!sheetName) {
      throw new Error('シート名が指定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log(`シート取得エラー (${sheetName}): ` + error.toString());
    throw error;
  }
}