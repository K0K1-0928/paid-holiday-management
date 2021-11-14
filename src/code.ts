const sheetId: string = 'SpreadSheetID';
const sheetName: string = '有給休暇管理表';
const eMailKey: string = 'メールアドレス';
const joiningDayKey: string = '入社日';
const thisYearGrantKey: string = '今年度付与日数';
const thisYearRemainKey: string = '今年度分残日数';
const lastYearGrantKey: string = '前年度付与日数';
const lastYearRemainKey: string = '前年度分残日数';
const yearBeforeLastRemainKey: string = '前々年度未消化';

function main() {
  let today = new Date();
  let vals = getRecords();
  vals.forEach((val: {}) => {
    if (isExactlySixMonths(today, val[joiningDayKey])) {
      movingCells(val, lastYearRemainKey, yearBeforeLastRemainKey);
      movingCells(val, thisYearRemainKey, lastYearRemainKey);
      movingCells(val, thisYearGrantKey, lastYearGrantKey);
      grantPaidHolidays(val, today);
    }
  });
}

/**
 * スプレッドシートのデータを、JSON配列形式で取得する
 *
 * @return {Array}  {{ key: string }[]}
 */
const getRecords = (): { key: string }[] => {
  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(sheetName);
  let rng = sh.getDataRange();
  let vals = rng.getValues();
  let records = [];
  let keys = vals.shift();
  vals.forEach((val) => {
    let json = {};
    for (let i = 0; i < val.length; i++) {
      json[keys[i]] = val[i];
    }
    records.push(json);
  });
  return records;
};

/**
 * 今日が入社日から6ヶ月後かどうか判定します。
 *
 * @param {Date} today - 今日
 * @param {Date} joiningDay - 入社日
 * @return {*}  {boolean}
 */
const isExactlySixMonths = (today: Date, joiningDay: Date): boolean => {
  // 入社日の6ヶ月後の日付を取得する
  let sixMonthsLetterDay = new Date(joiningDay);
  sixMonthsLetterDay.setMonth(sixMonthsLetterDay.getMonth() + 6);

  // 今日の月・日と入社6ヶ月後の月・日を比較し、一致していればtrueを返す
  if (today.getMonth() !== sixMonthsLetterDay.getMonth()) return false;
  if (today.getDate() !== sixMonthsLetterDay.getDate()) return false;
  return true;
};

/**
 * 社員情報のメールアドレスをキーとして操作行を特定し、
 * 参照元セルの値を操作先セルに移動します。
 *
 * @param {object} value - 社員情報JSON
 * @param {string} reference - 参照元セルのプロパティ名
 * @param {string} target - 操作先セルのプロパティ名
 */
const movingCells = (value: any, reference: string, target: string): void => {
  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(sheetName);
  const keysRow = 1;

  // 社員情報のメールアドレスと一致する行を操作対象行番号として取得する
  let addressCol = findColumn(sh, eMailKey, keysRow);
  let userRow = findRow(sh, value[eMailKey], addressCol);

  // 操作先セルに参照元セルの値を入力する
  let targetCol = findColumn(sh, target, keysRow);
  sh.getRange(userRow, targetCol).setValue(value[reference]);

  // 参照元セルの値を削除する
  let referenceCol = findColumn(sh, reference, keysRow);
  sh.getRange(userRow, referenceCol).clear();
};

/**
 * 社員情報のメールアドレスをキーとして操作行を特定し、
 * 勤続年数に応じて年次有給休暇の日数を「今年度付与日数」と「今年度分残日数」に記入します。
 *
 * @param {object} value - 社員情報JSON
 * @param {Date} today - 今日
 */
const grantPaidHolidays = (value: any, today: Date): void => {
  let sixMonthsLetterDay = new Date(value[joiningDayKey]);
  sixMonthsLetterDay.setMonth(sixMonthsLetterDay.getMonth() + 6);
  let lengthOfService = today.getFullYear() - sixMonthsLetterDay.getFullYear();
  let accrualPaidHolidays = 0;

  // 勤続年数に応じて年次有給休暇の日数を決定する
  if (lengthOfService == 0) {
    accrualPaidHolidays = 10;
  } else if (lengthOfService == 1) {
    accrualPaidHolidays = 11;
  } else if (lengthOfService == 2) {
    accrualPaidHolidays = 12;
  } else if (lengthOfService == 3) {
    accrualPaidHolidays = 14;
  } else if (lengthOfService == 4) {
    accrualPaidHolidays = 16;
  } else if (lengthOfService == 5) {
    accrualPaidHolidays = 18;
  } else if (lengthOfService >= 6) {
    accrualPaidHolidays = 20;
  }

  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(sheetName);
  const keysRow = 1;

  // 社員情報のメールアドレスと一致する行を操作対象として取得する
  let addressCol = findColumn(sh, eMailKey, keysRow);
  let userRow = findRow(sh, value[eMailKey], addressCol);
  let grantCol = findColumn(sh, thisYearGrantKey, keysRow);
  let remainCol = findColumn(sh, thisYearRemainKey, keysRow);

  // 年次有給休暇日数をシートに入力する
  sh.getRange(userRow, grantCol).setValue(accrualPaidHolidays);
  sh.getRange(userRow, remainCol).setValue(accrualPaidHolidays);
};

/**
 * 検索対象列から検索対象文字列と一致する行番号を返す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} value - 検索対象文字列
 * @param {number} column - 検索対象列の番号
 * @return {*}  {number}
 */
const findRow = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  value: string,
  column: number
): number => {
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][column - 1] === value) {
      return i + 1;
    }
  }
  return 0;
};

/**
 * 検索対象行から検索対象文字列と一致する列番号を返す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} value - 検索対象文字列
 * @param {number} row - 検索対象行の番号
 * @return {*}  {number}
 */
const findColumn = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  value: string,
  row: number
): number => {
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data[row - 1].length; i++) {
    if (data[row - 1][i] === value) {
      return i + 1;
    }
  }
  return 0;
};
