/* global Excel */

export function setRangeBold(context, rangeAddress) {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
  range.format.font.bold = true;
  return range;
}

export function clearRange(context, rangeAddress) {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
  range.clear();
  return range;
}

export function autofitColumns(context, usedRange) {
  usedRange.load("address"); // Load the address property
  return context.sync().then(function () {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(usedRange.address);
    range.format.autofitColumns();
    return range;
  });
}

export function setColumnWidth(context, columnsOrRange, width) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  if (Array.isArray(columnsOrRange)) {
    columnsOrRange.forEach(column => {
      sheet.getRange(column + ":" + column).format.columnWidth = width;
    });
  } else {
    sheet.getRange(columnsOrRange).format.columnWidth = width;
  }
}

export function setRangeCenter(context, rangeAddress) {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
  range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
  return range;
}

export function setRangeRight(context, rangeAddress) {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
  range.format.horizontalAlignment = Excel.HorizontalAlignment.right;
  return range;
}

export function setFontSize(context, rangeAddress, fontSize) {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
  range.format.font.size = fontSize;
  return range;
}

export function setRangeItalic(context, rangeAddress) {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
  range.format.font.italic = true;
  return range;
}
