function border() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("名簿");
  sheet
    .getRange(1, 1)
    .setValue(
      `=ARRAYFORMULA(IF(ISBLANK($B$1:$B$150),"",COUNTIFS($B$1:$B$150,"<>'",ROW($B$1:$B$150),"<="&ROW($B$1:$B$150)-1)))`
    );

  let range = sheet.getRange("A1:V150");
  range.setBorder(
    false,
    false,
    false,
    false,
    true,
    false,
    "black",
    SpreadsheetApp.BorderStyle.SOLID
  );
  range.setBorder(
    null,
    null,
    null,
    null,
    null,
    true,
    "black",
    SpreadsheetApp.BorderStyle.DASHED
  );

  range = sheet.getRange("B1:G150");
  range.setBorder(
    null,
    true,
    null,
    true,
    null,
    null,
    "black",
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  range = sheet.getRange("U1:U150");
  range.setBorder(
    null,
    null,
    null,
    true,
    null,
    null,
    "black",
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}
