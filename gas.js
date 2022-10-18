function copySheet() {
    const date = Utilities.formatDate(new Date(), "JST", "YYYY-MM");
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (spreadsheet.getSheetByName(date)) {
        console.log(date, "は既に存在します")
        return;
    }

    const main_sheet = spreadsheet.getSheetByName('名簿');
    const tmp_sheet = spreadsheet.getSheetByName('テンプレート');
    const copy_sheet = tmp_sheet.copyTo(spreadsheet).activate();

    copy_sheet.setName(date);
    spreadsheet.moveActiveSheet(2);

    const get_data = main_sheet.getDataRange().getValues();
    let add_data = []

    for (let i = 1; i < get_data.length; i++) {
        add_data.push([get_data[i][1], get_data[i][2], get_data[i][3], get_data[i][4], get_data[i][5], get_data[i][6]]);
    }
    const set_range = copy_sheet.getRange(2, 2, add_data.length + 0, add_data[0].length);
    set_range.setValues(add_data);

    console.log("--fin--");
}

function border() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('名簿');
    let cell = sheet.getRange("A1:H150");
    sheet.getRange(1, 1).setValue(`=ARRAYFORMULA(IF(ISBLANK($B$1:$B$150),"",COUNTIFS($B$1:$B$150,"<>'",ROW($B$1:$B$150),"<="&ROW($B$1:$B$150)-1)))`)
    sheet.getRange(1, 9).setValue(`=ARRAYFORMULA($D$1:$D$150 & $C$1:$C$150 & $E$1:$E$150 & IF(ISBLANK($G$1:$G$150),"", TEXT($F$1:$F$150,"00")))`)
    cell.setBorder(false, false, false, false, true, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    cell.setBorder(null, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.DASHED);

    cell = sheet.getRange("B1:G150");
    cell.setBorder(null, true, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function format() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('名簿');
    sheet.getRange(1, 1, 300, 9).setBackground('black');

    const sort_data = sheet.getRange(2, 1, 300, 9);
    const first_cell = sheet.getRange(1, 2);
    const first_cell_value = first_cell.getValue();
    first_cell.setValue(`ソートを開始します`);
    sort_data.sort({ column: 9, ascending: true });

    const lastRow = sheet.getLastRow();
    let del_lock = true;
    const cells = sheet.getRange(1, 2, lastRow, 2).getValues();

    first_cell.setFontColor('white');
    for (let i = lastRow; i > 1; i--) {
        first_cell.setValue(`あと ${i - 1} 行`);

        if (!cells[i - 1][0] & !del_lock) {
            sheet.deleteRow(i);
        } else if (cells[i - 1][0]) {
            del_lock = false;
        }
    }

    border(sheet);
    sheet.getRange(2, 1, 300, 1).clear();
    sheet.getRange(2, 9, 300, 9).clear();
    first_cell.setValue(first_cell_value);
    first_cell.setFontColor('black');
    sheet.getRange(1, 1, 300, 9).setBackground('white');
}
