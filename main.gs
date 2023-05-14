/**
 * @description 将指定Sheet指定范围内容和格式清空
 */
function clearQuerySheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Sheet1')
  var range = sheet.getRange("A2:D")

  range.clearContent() // 清空内容
  range.clearFormat() // 删除格式
}

/**
 * @description 删除传入数组的重复内容
 * @param array 
 * @returns {Array} uniqueArray
 */
function removeDuplicatesFromArray(array) {
  var uniqueArray = [];
  var seen = {};

  for (var i = 0; i < array.length; i++) {
    var row = array[i];
    var uniqueRow = [];

    for (var j = 0; j < row.length; j++) {
      var cellValue = row[j];

      if (!seen.hasOwnProperty(cellValue)) {
        seen[cellValue] = true;
        uniqueRow.push(cellValue);
      } else {
        uniqueRow.push("")
      }
    }

    uniqueArray.push(uniqueRow);
  }

  return uniqueArray;
}