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