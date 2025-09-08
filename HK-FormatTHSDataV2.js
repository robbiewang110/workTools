Attribute Module_Name = "Module11"
/*
###
同花顺过滤代码
市值；行业；PE；PB；2025.6.30净利润；最近5季度净利润；最近5季度净利润增长率；最近3年每年净利润增长率；2025.6.30营收；2025.6.30营收增长率；最近5季度营收；最近5季度营收增速；最近3年每年营收增速；最近5日涨幅；最近10日涨幅；本月涨幅；2025年涨幅；2024.1.1至今涨幅；2024.9.24至今涨幅；(股价/2024.10.8日股价); (股价/2024.1.1日以来最低股价);(股价/2024.1.1日以来最高股价);股价/5日均线；股价/20日均线；股价/60日均线；5日均线/10日均线；所属概念；业绩预告；业绩预告类型；港股通标的；

###
请帮我写一个WPS表格格式设置的JS代码，我的WPS版本是2025版本的，具体要求如下：
请选择名称为`选股结果`的表格进行操作;
1、其中将所有行的行高设置为17.5
2、其中1、2行都是标题，格式设置背景颜色为51,119,255，文字颜色为255,255,255；此外WPS JS API 使用的是 BGR 顺序而非 RGB 顺序，请注意做转换；
3、并将表格第2行标记为可以过滤；
4、将列 C-R,U,V-Z,AA-AG,AM-AS,AY-AZ列的格式设置为 `#,##0.00`，前面列设置为一个变量我可以方便改。
5、将列AF-AL,AT-AX,BB的格式设置为 "0!.00,,\"亿\";-0!.00,,\"亿\""，前面列设置为一个变量我可以方便改。
6、将列 D,V-Z,AA-AG,AM-AS设置条件格式规则，具体规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于20的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>20的时候背景为255,199,206，字体显示为191,71,78；
7、将列 S,AA-AG,AM-AS设置条件格式规则，具体规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于30的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>30的时候背景为255,199,206，字体显示为191,71,78；
8、对列 AH-AL设置条件格式规则，具体规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于1000000000的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>1000000000的时候背景为255,199,206，字体显示为191,71,78；此外WPS JS API 使用的是 BGR 顺序而非 RGB 顺序，请注意做转换；
9、对列AT-AX设置条件格式规则，具体规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于500000000的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>500000000的时候背景为255,199,206，字体显示为191,71,78；此外WPS JS API 使用的是 BGR 顺序而非 RGB 顺序，请注意做转换；
10、将第5步的所有列的列宽设置为6
11、将第4步的所有列的列宽设置为6
*/
function formatStockSelectionTable() {
    try {
        let app = Application;
        let sheet = app.Sheets.Item("选股结果");
        app.ScreenUpdating = false;
        
        // 获取最后一行和最后一列
        let lastRow = sheet.UsedRange.Rows.Count;
        let lastCol = sheet.UsedRange.Columns.Count;
        
        // 1. 设置所有行高为17.5
        sheet.Rows.RowHeight = 17.5;
        
        // 2. 设置标题行格式（使用BGR顺序）
        let titleRows = sheet.Range("1:2");
        titleRows.Interior.Color = RGB_BGR(51, 119, 255); // 背景色
        titleRows.Font.Color = RGB_BGR(255, 255, 255);    // 文字颜色
        titleRows.Font.Bold = true;
        
        // 3. 启用第2行过滤
        let filterRange = sheet.Range(sheet.Cells(2, 1), sheet.Cells(lastRow, lastCol));
        sheet.AutoFilterMode = false; // 清除现有筛选
        filterRange.AutoFilter(1);    // 启用筛选
        
        // 4. 设置数字格式列（可修改的变量）
        let numberFormatRanges = "C-R,U,V-Z,AA-AG,AM-AS,AY-AZ";
        let numberFormatCols = parseColumnRanges(numberFormatRanges);
        
        // 5. 设置自定义格式列（可修改的变量）
        let customFormatRanges = "AF-AL,AT-AX,BB";
        let customFormatCols = parseColumnRanges(customFormatRanges);
        
        // 设置数值格式和列宽
        setColumnsFormat(sheet, numberFormatCols, "#,##0.00", 6); // 第4步列 + 第11步列宽
        setColumnsFormat(sheet, customFormatCols, "0!.00,,\"亿\";-0!.00,,\"亿\"", 6); // 第5步列 + 第10步列宽
        
        // 6. 第一组条件格式列（规则：0-20）
        let cond1Ranges = "D,V-Z,AA-AG,AM-AS";
        let cond1Cols = parseColumnRanges(cond1Ranges);
        cond1Cols.forEach(col => {
            setConditionalFormatting(sheet, col, 3, lastRow, [
                {type: "less", value: 0, 
                 font: RGB_BGR(74, 148, 74), 
                 fill: RGB_BGR(198, 239, 206)},
                {type: "between", min: 0, max: 20, 
                 font: RGB_BGR(185, 141, 48), 
                 fill: RGB_BGR(255, 235, 156)},
                {type: "greaterEqual", value: 20, 
                 font: RGB_BGR(191, 71, 78), 
                 fill: RGB_BGR(255, 199, 206)}
            ]);
        });
        
        // 7. 第二组条件格式列（规则：0-30）
        let cond2Ranges = "S,AA-AG,AM-AS";
        let cond2Cols = parseColumnRanges(cond2Ranges);
        
        // 处理重复列（先清除条件格式）
        cond2Cols.forEach(col => {
            // 清除这些列上的条件格式（避免规则冲突）
            let range = sheet.Range(sheet.Cells(3, col), sheet.Cells(lastRow, col));
            range.FormatConditions.Delete();
            
            // 应用新规则
            setConditionalFormatting(sheet, col, 3, lastRow, [
                {type: "less", value: 0, 
                 font: RGB_BGR(74, 148, 74), 
                 fill: RGB_BGR(198, 239, 206)},
                {type: "between", min: 0, max: 30, 
                 font: RGB_BGR(185, 141, 48), 
                 fill: RGB_BGR(255, 235, 156)},
                {type: "greaterEqual", value: 30, 
                 font: RGB_BGR(191, 71, 78), 
                 fill: RGB_BGR(255, 199, 206)}
            ]);
        });
        
        // 8. 第三组条件格式列（大数值规则1）
        let cond3Ranges = "AH-AL";
        let cond3Cols = parseColumnRanges(cond3Ranges);
        cond3Cols.forEach(col => {
            setConditionalFormatting(sheet, col, 3, lastRow, [
                {type: "less", value: 0, 
                 font: RGB_BGR(74, 148, 74), 
                 fill: RGB_BGR(198, 239, 206)},
                {type: "between", min: 0, max: 1000000000, 
                 font: RGB_BGR(185, 141, 48), 
                 fill: RGB_BGR(255, 235, 156)},
                {type: "greaterEqual", value: 1000000000, 
                 font: RGB_BGR(191, 71, 78), 
                 fill: RGB_BGR(255, 199, 206)}
            ]);
        });
        
        // 9. 第四组条件格式列（大数值规则2）
        let cond4Ranges = "AT-AX";
        let cond4Cols = parseColumnRanges(cond4Ranges);
        cond4Cols.forEach(col => {
            setConditionalFormatting(sheet, col, 3, lastRow, [
                {type: "less", value: 0, 
                 font: RGB_BGR(74, 148, 74), 
                 fill: RGB_BGR(198, 239, 206)},
                {type: "between", min: 0, max: 500000000, 
                 font: RGB_BGR(185, 141, 48), 
                 fill: RGB_BGR(255, 235, 156)},
                {type: "greaterEqual", value: 500000000, 
                 font: RGB_BGR(191, 71, 78), 
                 fill: RGB_BGR(255, 199, 206)}
            ]);
        });
        
        app.ScreenUpdating = true;
        
        // 显示完成信息
        let msg = "选股结果表设置完成！\n" +
                 "✓ 所有行高: 17.5\n" +
                 "✓ 标题行背景: RGB(51,119,255), 文字: 白色\n" +
                 "✓ 数字格式列: " + numberFormatCols.length + "列 (列宽6)\n" +
                 "✓ 亿单位列: " + customFormatCols.length + "列 (列宽6)\n" +
                 "✓ 规则1条件格式列: " + cond1Cols.length + "列 (0-20规则)\n" +
                 "✓ 规则2条件格式列: " + cond2Cols.length + "列 (0-30规则)\n" +
                 "✓ 规则3条件格式列: " + cond3Cols.length + "列 (10亿规则)\n" +
                 "✓ 规则4条件格式列: " + cond4Cols.length + "列 (5亿规则)\n" +
                 "✓ 按第二行设置筛选";
        alert(msg);
        
    } catch (error) {
        Application.ScreenUpdating = true;
        alert("操作失败: " + error.message + 
              "\n可能原因：\n1. 工作表不存在\n2. 工作表受保护\n3. 无效的列范围");
    }
}

// ===== 辅助函数 ===== //

// 设置列格式和宽度
function setColumnsFormat(sheet, columns, format, width) {
    columns.forEach(col => {
        let colRange = sheet.Columns(col);
        colRange.NumberFormat = format;
        colRange.ColumnWidth = width;
    });
}

// 解析列范围字符串
function parseColumnRanges(rangeStr) {
    let ranges = rangeStr.split(',');
    let columns = [];
    
    ranges.forEach(r => {
        r = r.trim();
        if (r.includes('-')) {
            let [start, end] = r.split('-');
            let startCol = getColumnNumber(start);
            let endCol = getColumnNumber(end);
            
            for (let col = startCol; col <= endCol; col++) {
                columns.push(getColumnLetter(col));
            }
        } else {
            columns.push(r);
        }
    });
    
    return columns;
}

// 设置条件格式
function setConditionalFormatting(sheet, col, startRow, lastRow, rules) {
    try {
        let dataRange = sheet.Range(sheet.Cells(startRow, col), sheet.Cells(lastRow, col));
        
        rules.forEach(rule => {
            let cond;
            if (rule.type === "less") {
                cond = dataRange.FormatConditions.Add(1, 6, rule.value); // xlCellValue, xlLess
            } else if (rule.type === "greaterEqual") {
                cond = dataRange.FormatConditions.Add(1, 7, rule.value); // xlCellValue, xlGreaterEqual
            } else if (rule.type === "between") {
                // 使用公式确保精确匹配
                let formula = `=AND(INDIRECT(ADDRESS(ROW(),COLUMN()))>${rule.min}, INDIRECT(ADDRESS(ROW(),COLUMN()))<${rule.max})`;
                cond = dataRange.FormatConditions.Add(2, 0, formula); // xlExpression
            }
            
            if (cond) {
                cond.Font.Color = rule.font;
                cond.Interior.Color = rule.fill;
            }
        });
    } catch (error) {
        console.error("设置条件格式出错: " + error.message + " (列" + col + ")");
    }
}

// 将列字母转换为列号
function getColumnNumber(columnLetter) {
    columnLetter = columnLetter.toUpperCase();
    let columnNumber = 0;
    
    for (let i = 0; i < columnLetter.length; i++) {
        let charCode = columnLetter.charCodeAt(i);
        if (charCode >= 65 && charCode <= 90) {
            columnNumber = columnNumber * 26 + (charCode - 64);
        }
    }
    
    return columnNumber;
}

// 将列号转换为列字母
function getColumnLetter(columnNumber) {
    let columnLetter = "";
    while (columnNumber > 0) {
        let modulo = (columnNumber - 1) % 26;
        columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
        columnNumber = Math.floor((columnNumber - modulo) / 26);
    }
    return columnLetter;
}

// 创建BGR颜色值 (WPS使用BGR顺序)
function RGB_BGR(r, g, b) {
    return (b << 16) | (g << 8) | r;
}