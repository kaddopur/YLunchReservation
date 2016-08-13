function parseResponse() {
  var RESPONSES_SHEET_NAME = '預約表單';
  var MEAL_SHEET_NAME = '主餐結果';
  var SALAD_SHEET_NAME = '沙拉結果';
  var MEAL_ORDER_REGEXP = new RegExp('主餐');
  var SALAD_ORDER_REGEXP = new RegExp('沙拉');
  var MEAL_LIMIT_COUNT = 120;
  var SALAD_LIMIT_COUNT = 120;
  var MEAL_OUTPUT_RANGE = 'A1:C' + MEAL_LIMIT_COUNT;
  var SALAD_OUTPUT_RANGE = 'A1:C' + SALAD_LIMIT_COUNT;

  var thisSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var responsesSheet = thisSpreadSheet.getSheetByName(RESPONSES_SHEET_NAME);
  var mealResultSheet = thisSpreadSheet.getSheetByName(MEAL_SHEET_NAME);
  var saladResultSheet = thisSpreadSheet.getSheetByName(SALAD_SHEET_NAME);
  var mealDataRange = mealResultSheet.getRange(MEAL_OUTPUT_RANGE);
  var saladDataRange = saladResultSheet.getRange(SALAD_OUTPUT_RANGE);

  var responsesDataRange = responsesSheet.getDataRange();
  var responsesDataValues = responsesDataRange.getValues();

  var mealResultArray = [];
  var saladResultArray = [];

  // parse sheet
  responsesDataValues.forEach(function(data, index) {
    // escape first row
    if (index === 0) {
      return;
    }

    // data[0] is timestamp
    var user = data[1];
    var number = data[2];
    var order = data[3];

    if (MEAL_ORDER_REGEXP.test(order)) {
      mealResultArray.push([number, user]);
    }
    if (SALAD_ORDER_REGEXP.test(order)) {
      saladResultArray.push([number, user]);
    }
  });

  // transform output data, fill empty string if user less than limit amount
  for (var i = 0; i < MEAL_LIMIT_COUNT; i++) {
    if (i >= mealResultArray.length) {
      mealResultArray[i] = [i + 1].concat(['', '']);
      continue;
    }

    mealResultArray[i] = [i + 1].concat(mealResultArray[i]);
  };
  for (var i = 0; i < SALAD_LIMIT_COUNT; i++) {
    if (i >= saladResultArray.length) {
      saladResultArray[i] = [i + 1].concat(['', '']);
      continue;
    }
    saladResultArray[i] = [i + 1].concat(saladResultArray[i]);
  }

  // output data
  mealDataRange.setValues(mealResultArray);
  saladDataRange.setValues(saladResultArray);
}
