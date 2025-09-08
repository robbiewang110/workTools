///////////////////////////////////////////////
//V9
//最后更新时间：20250907--01:00:00


// 辅助函数：创建BGR颜色值 (WPS使用BGR顺序)
function RGB_BGR(r, g, b) {
    return (b << 16) | (g << 8) | r;
}

// 辅助函数：解析列范围字符串（如"N-R"或"X-AD,AE-AK"）
function parseColumnRange(rangeStr) {
    let result = [];
    let parts = rangeStr.split(",");

    for (let part of parts) {
        part = part.trim();
        if (part.includes("-")) {
            let [start, end] = part.split("-");
            let startNum = columnLetterToNumber(start);
            let endNum = columnLetterToNumber(end);

            for (let i = startNum; i <= endNum; i++) {
                result.push(getColumnLetter(i));
            }
        } else {
            result.push(part);
        }
    }

    return result;
}

// 辅助函数：设置条件格式
function setConditionalFormatting(sheet, col, startRow, lastRow, rules) {
    try {
        let dataRange = sheet.Range(
            sheet.Cells(startRow, col),
            sheet.Cells(lastRow, col)
        );
        dataRange.FormatConditions.Delete();

        rules.forEach((rule) => {
            let cond;
            if (rule.type === "less") {
                cond = dataRange.FormatConditions.Add(1, 6, rule.value); // xlCellValue, xlLess
            } else if (rule.type === "greaterEqual") {
                cond = dataRange.FormatConditions.Add(1, 7, rule.value); // xlCellValue, xlGreaterEqual
            } else if (rule.type === "between") {
                // 使用公式确保精确匹配
                let formula = `=AND(INDIRECT(ADDRESS(ROW(), COLUMN()))>${rule.min}, INDIRECT(ADDRESS(ROW(), COLUMN()))<${rule.max})`;
                cond = dataRange.FormatConditions.Add(2, 0, formula); // xlExpression
            }

            if (cond) {
                cond.Font.Color = rule.font;
                cond.Interior.Color = rule.fill;
            }
        });
    } catch (error) {
        console.error("设置条件格式出错: " + error.message);
    }
}


// 删除底部合并行
function deleteLastMergedRow(sheet) {
    let lastRow = sheet.UsedRange.Rows.Count;

    // 检查最后一行是否有合并单元格
    let hasMergedCells = false;

    // 遍历最后一行的所有单元格
    for (let col = 1; col <= sheet.UsedRange.Columns.Count; col++) {
        let cell = sheet.Cells(lastRow, col);
        if (cell.MergeCells) {
            hasMergedCells = true;
            break;
        }
    }

    // 如果有合并单元格，则删除整行
    if (hasMergedCells) {
        sheet.Rows(lastRow).Delete();
        return true;
    }

    return false;
}

// 调试函数：显示第一行所有单元格的值
function debugFirstRowValues(sheet) {
    let lastCol = sheet.UsedRange.Columns.Count;
    let debugInfo = "第一行单元格值:\n\n";

    for (let col = 1; col <= lastCol; col++) {
        let cell = sheet.Cells(1, col);
        let cellValue = getCellValue(cell);
        let isMerged = cell.MergeCells ? "是" : "否";

        debugInfo += `列 ${col}: "${cellValue}" (合并: ${isMerged})\n`;

        if (cell.MergeCells) {
            try {
                let topLeftCell = getMergeAreaTopLeftCell(sheet, 1, col);
                let topLeftValue = getCellValue(topLeftCell);
                debugInfo += `  左上角值: "${topLeftValue}"\n`;
            } catch (error) {
                debugInfo += `  获取左上角值失败: ${error.message}\n`;
            }
        }
    }

    alert(debugInfo);
}

// 获取单元格值

function getCellValue(cell) {
    try {
        if (cell.Value !== undefined && cell.Value !== null) {
            if (typeof cell.Value === "function") {
                return cell.Value().toString().trim();
            }
            return cell.Value.toString().trim();
        }
        return null;
    } catch (error) {
        try {
            if (cell.Formula) {
                return cell.Formula.toString().trim();
            }
            if (cell.Text) {
                return cell.Text.toString().trim();
            }
            return null;
        } catch (e) {
            return null;
        }
    }
}

// 获取合并区域的左上角单元格
function getMergeAreaTopLeftCell(sheet, row, col) {
    let cell = sheet.Cells(row, col);

    if (cell.MergeCells) {
        try {
            let mergeArea = cell.MergeArea;
            let topLeftRow = mergeArea.Row;
            let topLeftCol = mergeArea.Column;
            return sheet.Cells(topLeftRow, topLeftCol);
        } catch (error) {
            console.error("获取合并区域失败: " + error.message);
            return cell;
        }
    }

    return cell;
}

// 辅助函数：将列号转换为列字母（1→A, 26→Z, 27→AA等）
function getColumnLetter(columnNumber) {
    let columnLetter = "";
    while (columnNumber > 0) {
        let modulo = (columnNumber - 1) % 26;
        columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
        columnNumber = Math.floor((columnNumber - modulo) / 26);
    }
    return columnLetter;
}

/**
 * 将列字母转换为列索引（A=1, B=2, ...）
 */
function getColumnIndex(columnLetter) {
    let index = 0;
    for (let i = 0; i < columnLetter.length; i++) {
        index = index * 26 + (columnLetter.charCodeAt(i) - "A".charCodeAt(0) + 1);
    }
    return index;
}

// 辅助函数：将列字母转换为列号（A→1, Z→26, AA→27等）
function columnLetterToNumber(columnLetter) {
    let columnNumber = 0;
    let length = columnLetter.length;

    for (let i = 0; i < length; i++) {
        columnNumber +=
            (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }

    return columnNumber;
}

// 查找包含指定标题的列
function findColumnByHeader(sheet, headerText) {
    let lastCol = sheet.UsedRange.Columns.Count;

    for (let col = 1; col <= lastCol; col++) {
        let cell = sheet.Cells(1, col);
        let cellValue = "";

        if (cell.MergeCells) {
            try {
                let topLeftCell = getMergeAreaTopLeftCell(sheet, 1, col);
                cellValue = getCellValue(topLeftCell);

                if (cellValue === headerText) {
                    return topLeftCell.Column;
                }
            } catch (error) {
                console.error("处理合并单元格时出错: " + error.message);
                continue;
            }
        } else {
            cellValue = getCellValue(cell);

            if (cellValue === headerText) {
                return col;
            }
        }
    }

    return -1;
}


/**
 * 计算已应用的过滤条件数量
 */
function countAppliedFilters(criteria) {
    let count = 0;
    if (criteria.industry) count++;
    // 计算概念条件数量
    if (criteria.concept) count++;
    if (criteria.concept2) count++;
    if (criteria.concept3) count++;
    // 计算市值条件数量
    if (criteria.marketValueMin !== null && !isNaN(criteria.marketValueMin))
        count++;
    if (criteria.marketValueMax !== null && !isNaN(criteria.marketValueMax))
        count++;
    return count;
}


// 复制单元格格式
function copyCellFormat(sourceCell, targetCell) {
    try {
        // 复制字体格式
        targetCell.Font.Name = sourceCell.Font.Name;
        targetCell.Font.Size = sourceCell.Font.Size;
        //targetCell.Font.Bold = sourceCell.Font.Bold;
        //targetCell.Font.Italic = sourceCell.Font.Italic;
        targetCell.Font.Color = sourceCell.Font.Color;
        //targetCell.Font.Underline = sourceCell.Font.Underline;

        // 复制背景格式
        //targetCell.Interior.Color = sourceCell.Interior.Color;
        //targetCell.Interior.Pattern = sourceCell.Interior.Pattern;

        // 复制边框格式
        /*let borderTypes = ["xlEdgeLeft", "xlEdgeTop", "xlEdgeBottom", "xlEdgeRight", "xlInsideVertical", "xlInsideHorizontal"];
            for (let i = 0; i < borderTypes.length; i++) {
                let borderType = borderTypes[i];
                targetCell.Borders(borderType).LineStyle = sourceCell.Borders(borderType).LineStyle;
                targetCell.Borders(borderType).Weight = sourceCell.Borders(borderType).Weight;
                targetCell.Borders(borderType).Color = sourceCell.Borders(borderType).Color;
            }
            */

        // 复制对齐方式
        //targetCell.HorizontalAlignment = sourceCell.HorizontalAlignment;
        //targetCell.VerticalAlignment = sourceCell.VerticalAlignment;

        // 复制数字格式
        //targetCell.NumberFormat = sourceCell.NumberFormat;
    } catch (error) {
        console.error("复制单元格格式时出错: " + error.message);
    }
}

// 处理"所属概念"列
function processConceptColumn(sheet) {
    // 查找"所属概念"列
    let conceptCol = findColumnByHeader(sheet, "所属概念");
    if (conceptCol === -1) {
        debugFirstRowValues(sheet);
        return false;
    }

    // 获取最后一列的位置
    let lastCol = sheet.UsedRange.Columns.Count;

    // 目标列是最后一列的下一列
    let targetCol = lastCol + 1;

    // 获取数据行数
    let lastRow = sheet.UsedRange.Rows.Count;

    // 处理标题行的合并状态
    let sourceHeader = getMergeAreaTopLeftCell(sheet, 1, conceptCol);
    let isHeaderMerged =
        sheet.Cells(1, conceptCol).MergeCells ||
        sheet.Cells(2, conceptCol).MergeCells;

    // 如果原始列的前两行是合并的，也在目标列创建相同的合并
    if (isHeaderMerged) {
        let mergeRange = sheet.Range(
            sheet.Cells(1, targetCol),
            sheet.Cells(2, targetCol)
        );
        mergeRange.Merge();

        // 复制标题
        let targetHeader = sheet.Cells(1, targetCol);
        targetHeader.Value2 = sourceHeader.Value2;
        copyCellFormat(sourceHeader, targetHeader);
    } else {
        // 复制第一行标题
        let targetHeader = sheet.Cells(1, targetCol);
        targetHeader.Value2 = sourceHeader.Value2;
        copyCellFormat(sourceHeader, targetHeader);

        // 复制第二行标题
        let sourceHeader2 = sheet.Cells(2, conceptCol);
        let targetHeader2 = sheet.Cells(2, targetCol);
        targetHeader2.Value2 = sourceHeader2.Value2;
        copyCellFormat(sourceHeader2, targetHeader2);
    }

    // 复制数据（从第3行开始）
    for (let row = 3; row <= lastRow; row++) {
        let sourceCell = sheet.Cells(row, conceptCol);
        let targetCell = sheet.Cells(row, targetCol);

        // 复制值和格式
        targetCell.Value2 = sourceCell.Value2;
        copyCellFormat(sourceCell, targetCell);

        // 处理数据行的合并状态
        if (sourceCell.MergeCells) {
            try {
                let mergeArea = sourceCell.MergeArea;
                let mergeRows = mergeArea.Rows.Count;

                // 在目标列创建相同的合并
                if (mergeRows > 1) {
                    let targetMergeRange = sheet.Range(
                        sheet.Cells(row, targetCol),
                        sheet.Cells(row + mergeRows - 1, targetCol)
                    );
                    targetMergeRange.Merge();

                    // 跳过已合并的行
                    row += mergeRows - 1;
                }
            } catch (error) {
                console.error("处理数据行合并时出错: " + error.message);
            }
        }
    }

    // 删除原始列
    sheet.Columns(conceptCol).Delete();
    return true;
}

//////////////////////////////////////////////////////////////////

/**
 * 清除过滤状态（修正版）
 * 按列单独判断合并状态，合并列用第1行过滤，非合并列用第2行过滤
 */
function ClearFilterSetting() {
    try {
        // 获取当前工作簿
        let wb = Application.Workbooks.Item(1);
        if (!wb) {
            alert("请确保有打开的工作簿");
            return;
        }

        // 获取所需工作表
        let wsResult = wb.Worksheets.Item("选股结果");
        if (!wsResult) {
            alert("请确保存在名为'选股结果'的工作表");
            return;
        }

        // 先移除已有的筛选
        if (wsResult.AutoFilterMode) {
            wsResult.AutoFilterMode = false;
        }

        //先取消所有行在隐藏：
        wsResult.UsedRange.EntireRow.Hidden = false;

        // 优化筛选范围定义 - 使用表头行而非整个数据区域
        let startCol = 1;
        let endCol = wsResult.UsedRange.Columns.Count;
        let headerRow = 2; // 整体以第2行为筛选基准
        let filterRange = wsResult.Range(
            wsResult.Cells(headerRow, startCol),
            wsResult.Cells(headerRow, endCol)
        );

        // 应用筛选（初始不设置任何条件）
        filterRange.AutoFilter();

    } catch (e) {
        alert("执行过程中发生错误：" + e.message);
        console.error(e);
    }
}



/**
 * 校验表格第一行结构是否符合要求
 * @param {Object} sheet - 工作表对象
 * @returns {boolean} - 校验是否通过
 */
function validateFirstRowStructure(sheet) {
    try {
        // 定义第一行预期的结构（空字符串表示合并单元格）
        const expectedStructure = [
            "股票代码",           // 1
            "股票简称",           // 2
            "涨跌幅(%)",          // 3
            "所属同花顺行业",      // 4
            "市盈率(pe)",         // 5（包含日期）
            "总市值(元)",         // 6（包含日期）
            "区间涨跌幅:前复权(%)",// 7
            "",                   // 8（合并）
            "",                   // 9（合并）
            "",                   // 10（合并）
            "",                   // 11（合并）
            "",                   // 12（合并）
            "市净率",             // 13（包含日期）
            "营业收入(元)",       // 14
            "",                   // 15（合并）
            "",                   // 16（合并）
            "",                   // 17（合并）
            "",                   // 18（合并）
            "净利润(元)",         // 19
            "",                   // 20（合并）
            "",                   // 21（合并）
            "",                   // 22（合并）
            "",                   // 23（合并）
            "营业收入(同比增长率)(%)",// 24
            "",                   // 25（合并）
            "",                   // 26（合并）
            "",                   // 27（合并）
            "",                   // 28（合并）
            "",                   // 29（合并）
            "",                   // 30（合并）
            "净利润同比增长率(%)",  // 31
            "",                   // 32（合并）
            "",                   // 33（合并）
            "",                   // 34（合并）
            "",                   // 35（合并）
            "",                   // 36（合并）
            "",                   // 37（合并）
            "收盘价:不复权(元)",   // 38
            "",                   // 39（合并）
            "[1] / [2]",          // 40
            "区间最低价:前复权(元) [3]",// 41（包含日期范围）
            "区间最低价:前复权日",   // 42（包含日期范围）
            "[1] / [3]",          // 43
            "区间最高价:前复权(元) [4]",// 44（包含日期范围）
            "区间最高价:前复权日",   // 45（包含日期范围）
            "[1] / [4]",          // 46
            "5日均线 [5]",        // 47（包含日期）
            "20日均线 [6]",       // 48（包含日期）
            "60日均线 [7]",       // 49（包含日期）
            "10日均线 [8]",       // 50（包含日期）
            "[1] / [5]",          // 51
            "[1] / [6]",          // 52
            "[1] / [7]",          // 53
            "[5] / [8]",          // 54
            "区间主力资金流向(元)",  // 55
            "",                   // 56（合并）
            "",                   // 57（合并）
            "预告净利润中值(元)",   // 58（包含日期）
            "业绩预告类型",       // 59（包含日期）
            "所属概念"            // 60
        ];

        // 存储错误信息
        const errors = [];
        // 获取第一行使用的列数
        const usedCols = sheet.UsedRange.Columns.Count;

        // 检查列数是否至少为59列
        if (usedCols < expectedStructure.length) {
            errors.push(`表格列数不足，至少需要${expectedStructure.length}列，当前只有${usedCols}列`);
        }

        // 遍历每一列检查内容
        for (let col = 1; col <= expectedStructure.length; col++) {
            const expected = expectedStructure[col - 1];
            const cell = getMergeAreaTopLeftCell(sheet, 1, col);
            const actualValue = getCellValue(cell) || "";

            // 检查合并状态：预期为空的列应该是合并单元格
            if (expected === "") {
                if (!cell.MergeCells) {
                    errors.push(`第${col}列：预期为合并单元格，但未合并`);
                }
                continue;
            }

            // 检查非空列的内容
            if (expected.includes("(pe)") ||
                expected.includes("总市值") ||
                expected.includes("市净率") ||
                expected.includes("日均线") ||
                expected.includes("预告净利润中值(元)") ||
                expected.includes("业绩预告类型")) {
                // 包含日期的单元格：检查前缀是否匹配
                if (!actualValue.startsWith(expected)) {
                    errors.push(`第${col}列：内容不符合要求，应为"${expected}日期"，实际为"${actualValue}"`);
                }
            } else if (expected.includes("区间最低价") ||
                expected.includes("区间最高价")) {
                // 包含日期范围的单元格：检查前缀是否匹配
                if (!actualValue.startsWith(expected)) {
                    errors.push(`第${col}列：内容不符合要求，应为"${expected}日期范围"，实际为"${actualValue}"`);
                }
            } else if (actualValue !== expected) {
                // 精确匹配的单元格
                errors.push(`第${col}列：内容不匹配，预期"${expected}"，实际"${actualValue}"`);
            }

            // 检查非空列是否不应合并
            //if (expected !== "" && cell.MergeCells) {
            //    errors.push(`第${col}列：不应为合并单元格，但被合并了`);
            //}
        }

        // 显示错误信息
        if (errors.length > 0) {
            alert(`表格第一行结构校验失败：\n\n${errors.join("\n")}`);
            return false;
        }

        return true;
    } catch (error) {
        alert(`校验第一行结构时出错：${error.message}`);
        return false;
    }
}

// 在主处理函数中添加校验调用（示例）
function ProcessStockSelectionTable() {
    try {
        let app = Application;
        let sheet = app.Sheets.Item("选股结果");
        app.ScreenUpdating = false;


        // 判断工作表是否处于保护状态
        if (sheet.ProtectContents) {
            sheet.Unprotect(); // 取消工作表保护
        }

        // 后续原有逻辑...
        // 1. 删除底部合并行
        let deleted = deleteLastMergedRow(sheet);
        if (deleted) {
            console.log("已删除表格底部空白合并行");
        }

        // 2. 处理"所属概念"列
        let conceptProcessed = processConceptColumn(sheet);
        if (conceptProcessed) {
            console.log(
                "已成功处理'所属概念'列：\n- 复制数据到最后一列右侧\n- 删除原始列"
            );
        }

        //做完上面的处理后校验表格格式是否有变化
        // 新增：校验第一行结构
        if (!validateFirstRowStructure(sheet)) {
            app.ScreenUpdating = true;
            alert("这个表格也预期的表格格式发生了变化，请手动检查！！");
            return; // 校验失败则终止处理
        } else {
            console.log("这个表格的格式检查正确。。将继续！！");
        }

        alert("对表格自动处理成功，请继续！！");

        app.ScreenUpdating = true;
    } catch (error) {
        Application.ScreenUpdating = true;
        alert(
            "操作失败: " +
            error.message +
            "\n可能原因：\n1. '选股结果'工作表不存在\n2. 工作表受保护\n3. 无效的列范围"
        );
    }
}


//自动设置行业板块在规则

//////////////////////////////////////////////////////////
/*
请帮我编写一个WPS表格格式的JS代码，我的WPS版本是2025版本的，我在要求的后面会附带以前的代码供参考。
请选择名称为`选股结果`的表格进行操作，具体要求如下：
1、定位的表格的最下面一行，如果这一行有单元格是合并了的，那么将这一行整个删除掉。
2、找到第一行的单元格内容为"所属概念"的列，记录下这一列的位置，将列命名为：原始列；定位到表格最后一列有数据的位置，将它的下一列设置为目标列；将原始列的所有数据都拷贝到目标列。最后将原始列删除掉。
3、在查找"所属概念"的列的时候要注意：第1、2行有的单元格有合并，合并形式是有的是多个第一行的合并，有的是1、2行的有合并。
4、在处理复制的时候要注意：我原始列中前面两行是有合并的，目的列的前面两行没有合并，这个可能会报错。
*/
function ProcessStockSelectionTable2() {
    try {
        let app = Application;
        let sheet = app.Sheets.Item("选股结果");
        app.ScreenUpdating = false;

        // 判断工作表是否处于保护状态
        if (sheet.ProtectContents) {
            sheet.Unprotect(); // 取消工作表保护
        }

        // 1. 删除底部合并行
        let deleted = deleteLastMergedRow(sheet);
        if (deleted) {
            alert("已删除表格底部空白合并行");
        } else {
            //alert("底部行无合并单元格，无需删除");
        }

        // 2. 处理"所属概念"列
        let conceptProcessed = processConceptColumn(sheet);
        if (conceptProcessed) {
            alert(
                "已成功处理'所属概念'列：\n- 复制数据到最后一列右侧\n- 删除原始列"
            );
        } else {
            //alert("未找到'所属概念'列，跳过此步骤");
        }

        app.ScreenUpdating = true;
        //alert("选股结果表处理完成！");
    } catch (error) {
        Application.ScreenUpdating = true;
        alert(
            "操作失败: " +
            error.message +
            "\n可能原因：\n1. '选股结果'工作表不存在\n2. 工作表受保护\n3. 无效的列范围"
        );
    }
}





/*#########################
#V5
请帮我编写一个WPS表格格式的JS代码，我的WPS版本是2025版本的，我在要求的后面会附带以前的代码供参考。
请选择名称为`选股结果`的表格进行操作，具体要求如下：
1、其中将所有行的行高设置为17.5；
2、其中1、2行都是标题，格式设置背景颜色为51,119,255，文字颜色为255,255,255；此外WPS JS API 使用的是 BGR 顺序而非 RGB 顺序，请注意做转换；
3、并将表格第2行标记为可以过滤；
4、下面我定义一些列，在后面用名字引用他们，请注意这些列可能会有交叉的，具体为：营业收入的列为：N-R；营业利润的列为：S-W；营收增长比例的列为：X-AD,AE-AK；市值的列为：F；主力资金的列为：BC-BE；预告利润列为：BF；当日涨跌幅为C列；区间涨跌幅的列为：C,G-L；默认格式列为：C,E-BF
5、对第4步中的默认格式列设置格式为 `#,##0.00`，并且将列宽设置为6；
6、将第4步中定义的 营业收入、营业利润、市值、主力资金、预告利润这些列的格式设置为"0!.00,,\"亿\";-0!.00,,\"亿\""
7、将第4步中定义的 区间涨跌幅 列设置条件格式，规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于20的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>20的时候背景为23,0,0，字体显示为255,255,255；
8、将第4步中定义的 当日涨跌幅 列设置条件格式，规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于8的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>8的时候背景为255,199,206，字体显示为191,71,78；
9、将第4步中定义的 营收增长比例 列设置条件格式，规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于30的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>30的时候背景为235,0,0，字体显示为255,255,255；
10、将第4步中定义的 营业收入 列设置条件格式，规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于1000000000的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>1000000000的时候背景为255,199,206，字体显示为191,71,78；此外WPS JS API 使用的是 BGR 顺序而非 RGB 顺序，请注意做转换；
11、将第4步中定义的 营业利润 列设置条件格式，规则为：数值<0的时候背景颜色为0,128,0，字体颜色为255,255,255；数值>0 且数值小于500000000的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>500000000的时候背景为255,199,206，字体显示为191,71,78；此外WPS JS API 使用的是 BGR 顺序而非 RGB 顺序，请注意做转换；
12、将第4步中定义的 主力资金 列设置条件格式，规则为：数值<0的时候背景颜色为198,239,206，字体颜色为74，148，74；数值>0 且数值小于50000000的时候显示背景颜色为255,235,156，字体颜色为185,141,48；数值>50000000的时候背景为255,199,206，字体显示为191,71,78；此外WPS JS API 使用的是 BGR 顺序而非 RGB 顺序，请注意做转换；
13、将第4步中定义的 营业利润、营业收入 列设置列宽为7；
14、注意将颜色也提取为变量方便后面进行配置；
*/
function FormatStockSelectionTable() {
    // 定义颜色常量（使用BGR顺序）
    const COLORS = {
        // 标题颜色
        TITLE_BG: RGB_BGR(51, 119, 255),
        TITLE_TEXT: RGB_BGR(255, 255, 255),

        // 条件格式颜色 - 绿色系
        GREEN_FILL: RGB_BGR(198, 239, 206),
        GREEN_TEXT: RGB_BGR(74, 148, 74),
        DARK_GREEN_FILL: RGB_BGR(0, 128, 0),

        // 条件格式颜色 - 黄色系
        YELLOW_FILL: RGB_BGR(255, 235, 156),
        YELLOW_TEXT: RGB_BGR(185, 141, 48),

        // 条件格式颜色 - 红色系
        RED_FILL: RGB_BGR(255, 199, 206),
        RED_TEXT: RGB_BGR(191, 71, 78),
        DARK_RED_FILL: RGB_BGR(235, 0, 0),

        // 白色文本
        WHITE_TEXT: RGB_BGR(255, 255, 255),
    };

    try {
        const app = Application;
        const sheet = app.Sheets.Item("选股结果");
        if (!sheet) {
            alert("未找到名为'选股结果'的工作表");
            return;
        }

        app.ScreenUpdating = false;
        const lastRow = sheet.UsedRange.Rows.Count;
        const lastCol = sheet.UsedRange.Columns.Count;

        // 存储统计信息
        const stats = {
            defaultFormatCols: 0,
            billionCols: 0,
            rangeChangeCols: 0,
            dailyChangeCol: 0,
            growthCols: 0,
            revenueCols: 0,
            profitCols: 0,
            mainFundCols: 0
        };

        // 1. 设置所有行的行高为17.5
        sheet.Rows.RowHeight = 17.5;

        // 2. 设置标题行样式
        const titleRows = sheet.Range("1:2");
        titleRows.Interior.Color = COLORS.TITLE_BG;
        titleRows.Font.Color = COLORS.TITLE_TEXT;
        titleRows.Font.Bold = true;

        // 3. 设置表格筛选
        sheet.AutoFilterMode = false;
        const filterRange = sheet.Range(sheet.Cells(2, 1), sheet.Cells(lastRow, lastCol));
        filterRange.AutoFilter(1);

        // 4. 解析列范围并缓存
        const columnRanges = {
            revenue: parseColumnRange("N-R"), // 营业收入
            profit: parseColumnRange("S-W"),// 营业利润
            growth: parseColumnRange("X-AD,AE-AK"),// 营收增长比例
            marketValue: parseColumnRange("F"),// 市值
            mainFund: parseColumnRange("BC-BE"),// 主力资金
            forecastProfit: parseColumnRange("BF"),// 预告利润
            dailyChange: parseColumnRange("C"),// 当日涨跌幅
            rangeChange: parseColumnRange("C,G-L"),// 区间涨跌幅
            defaultFormat: parseColumnRange("C,E-BF")// 默认格式列
        };

        // 更新统计数据
        Object.keys(stats).forEach(key => {
            stats[key] = columnRanges[key.replace('Cols', '')]?.length || 0;
        });

        // 5. 设置默认格式列
        applyColumnFormat(columnRanges.defaultFormat, {
            numberFormat: "#,##0.00",
            columnWidth: 6
        }, sheet);

        // 6. 设置亿单位列的格式
        const billionCols = [
            ...columnRanges.revenue,
            ...columnRanges.profit,
            ...columnRanges.marketValue,
            ...columnRanges.mainFund,
            ...columnRanges.forecastProfit
        ];
        applyColumnFormat(billionCols, {
            numberFormat: '0!.00,,"亿";-0!.00,,"亿"'
        }, sheet);

        // 7. 设置区间涨跌幅的条件格式
        applyConditionalFormattingToColumns(
            columnRanges.rangeChange,
            sheet,
            3,
            lastRow,
            [
                { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                { type: "between", min: 0, max: 20, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                { type: "greaterEqual", value: 20, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_RED_FILL }
            ]
        );

        // 8. 设置当日涨跌幅的条件格式
        applyConditionalFormattingToColumns(
            columnRanges.dailyChange,
            sheet,
            3,
            lastRow,
            [
                { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                { type: "between", min: 0, max: 8, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                { type: "greaterEqual", value: 8, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
            ]
        );

        // 9. 设置营收增长比例的条件格式
        applyConditionalFormattingToColumns(
            columnRanges.growth,
            sheet,
            3,
            lastRow,
            [
                { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                { type: "between", min: 0, max: 30, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                { type: "greaterEqual", value: 30, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_RED_FILL }
            ]
        );

        // 10. 设置营业收入的条件格式
        applyConditionalFormattingToColumns(
            columnRanges.revenue,
            sheet,
            3,
            lastRow,
            [
                { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                { type: "between", min: 0, max: 1000000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                { type: "greaterEqual", value: 1000000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
            ]
        );

        // 11. 设置营业利润的条件格式
        applyConditionalFormattingToColumns(
            columnRanges.profit,
            sheet,
            3,
            lastRow,
            [
                { type: "less", value: 0, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_GREEN_FILL },
                { type: "between", min: 0, max: 500000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                { type: "greaterEqual", value: 500000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
            ]
        );

        // 12. 设置主力资金的条件格式
        applyConditionalFormattingToColumns(
            columnRanges.mainFund,
            sheet,
            3,
            lastRow,
            [
                { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                { type: "between", min: 0, max: 50000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                { type: "greaterEqual", value: 50000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
            ]
        );

        // 13. 设置营业收入和营业利润的列宽为7
        const width7Cols = [...columnRanges.revenue, ...columnRanges.profit];
        applyColumnFormat(width7Cols, { columnWidth: 7 }, sheet);

        app.ScreenUpdating = true;

        // 显示完成信息
        alert(`选股结果表设置完成！
✓ 所有行高: 17.5
✓ 标题行背景: RGB(51,119,255), 文字: 白色
✓ 默认格式列: ${stats.defaultFormatCols}列 (列宽6)
✓ 亿单位列: ${stats.billionCols}列
✓ 区间涨跌幅条件格式列: ${stats.rangeChangeCols}列
✓ 当日涨跌幅条件格式列: ${stats.dailyChangeCol}列
✓ 营收增长比例条件格式列: ${stats.growthCols}列
✓ 营业收入条件格式列: ${stats.revenueCols}列
✓ 营业利润条件格式列: ${stats.profitCols}列
✓ 主力资金条件格式列: ${stats.mainFundCols}列
✓ 营业收入/利润列宽: 7
✓ 按第二行设置筛选`);

    } catch (error) {
        Application.ScreenUpdating = true;
        alert(`操作失败: ${error.message}
可能原因：
1. '选股结果'工作表不存在
2. 工作表受保护
3. 无效的列范围`);
    }
}

//总的调用入口：
function AutomaticProcesMain() {
    //自动调整表格：
    ProcessStockSelectionTable();

    //自动处理表格的显示格式
    FormatStockSelectionTable();
}

////////////////###################////////////
//豆包优化后的：
/**
 * 在概念板块表中设置引用原始数据表的公式
 * 先检查概念板块是否存在，不存在不存在则创建
 */
function createConceptTableWithFormulas() {
    try {
        // 性能优化：提前获取应用实例
        const app = Application;
        // 获取当前工作簿
        let wb = app.Workbooks.Item(1);
        if (!wb) {
            alert("请确保有打开的工作簿");
            return;
        }

        // 定义表名
        const targetSheetName = "概念板块";
        const sourceSheetName = "选股结果";

        // 检查原始数据表是否存在
        let sourceSheet;
        try {
            sourceSheet = wb.Worksheets.Item(sourceSheetName);
        } catch (e) {
            alert(`未找到"${sourceSheetName}"工作表，请先创建该表`);
            return;
        }

        // 检查概念板块是否存在，不存在则创建
        let targetSheet;
        try {
            targetSheet = wb.Worksheets.Item(targetSheetName);
        } catch (e) {
            targetSheet = wb.Worksheets.Add();
            targetSheet.Name = targetSheetName;
        }

        // 性能优化：关闭自动计算、屏幕更新和事件
        const originalCalculation = app.Calculation;
        const originalScreenUpdating = app.ScreenUpdating;
        const originalEnableEvents = app.EnableEvents;
        app.Calculation = 1; // xlManual
        app.ScreenUpdating = false;
        app.EnableEvents = false;

        try {
            // 查找所属概念列的位置
            let tmpConceptCol = findColumnByHeader(sourceSheet, "所属概念");
            if (tmpConceptCol === -1) {
                debugFirstRowValues(sourceSheet); // 修复原代码中sheet未定义的问题
                return false;
            }
            const tmpConceptHeader = getColumnLetter(tmpConceptCol);

            // 获取实际数据范围，避免全列扫描
            const lastRow = sourceSheet.UsedRange.Rows.Count;
            const biRange = `'${sourceSheetName}'!${tmpConceptHeader}$2:${tmpConceptHeader}$${lastRow}`;
            const sourceRef = `'${sourceSheetName}'!`; // 缓存工作表引用字符串

            console.log("开始设置公式...");

            // 批量设置公式 - 使用数组一次性赋值提升性能
            const formulas = [
                ["B3", `=COUNTIF(${biRange},"*"&A3&"*")`],
                ["C3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}F:F)`],
                ["D3", `=C3/B3`],
                ["E3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}C:C)/B3`],//当日涨幅
                ["F3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}G:G)/B3`],//区间涨幅
                ["G3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}H:H)/B3`],
                ["H3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}I:I)/B3`],
                ["I3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}J:J)/B3`],
                ["J3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}K:K)/B3`],
                ["K3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}L:L)/B3`],
                ["L3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}BC:BC)/B3`],
                ["M3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}BD:BD)/B3`],
                ["N3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}BE:BE)/B3`],
                ["O3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}N:N)/B3`],
                ["P3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}O:O)/B3`],
                ["Q3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}P:P)/B3`],
                ["R3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}Q:Q)/B3`],
                ["S3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}R:R)/B3`],//净利润
                ["T3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}S:S)/B3`],
                ["U3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}T:T)/B3`],
                ["V3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}U:U)/B3`],
                ["W3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}V:V)/B3`],
                ["X3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}W:W)/B3`],//营收同比
                ["Y3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}X:X)/B3`],
                ["Z3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}Y:Y)/B3`],
                ["AA3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}Z:Z)/B3`],
                ["AB3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AA:AA)/B3`],
                ["AC3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AB:AB)/B3`],
                ["AD3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AC:AC)/B3`],
                ["AE3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AD:AD)/B3`],
                ["AF3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AF:AF)/B3`],
                ["AG3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AG:AG)/B3`],
                ["AH3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AH:AH)/B3`],
                ["AI3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AI:AI)/B3`],
                ["AJ3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AJ:AJ)/B3`],
                ["AK3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AK:AK)/B3`],
                ["AL3", `=SUMIF(${biRange},"*"&A3&"*",${sourceRef}AL:AL)/B3`]
            ];

            // 批量设置公式（减少与Excel交互次数）
            formulas.forEach(([cell, formula]) => {
                targetSheet.Range(cell).Formula = formula;
            });

            console.log("公式设置完成...");

            // 定义颜色常量（使用BGR顺序）
            const COLORS = {
                TITLE_BG: RGB_BGR(51, 119, 255),
                TITLE_TEXT: RGB_BGR(255, 255, 255),
                GREEN_FILL: RGB_BGR(198, 239, 206),
                GREEN_TEXT: RGB_BGR(74, 148, 74),
                DARK_GREEN_FILL: RGB_BGR(0, 128, 0),
                YELLOW_FILL: RGB_BGR(255, 235, 156),
                YELLOW_TEXT: RGB_BGR(185, 141, 48),
                RED_FILL: RGB_BGR(255, 199, 206),
                RED_TEXT: RGB_BGR(191, 71, 78),
                DARK_RED_FILL: RGB_BGR(235, 0, 0),
                WHITE_TEXT: RGB_BGR(255, 255, 255),
            };

            // 解析列范围（一次性解析所有范围）
            const revenueCols = parseColumnRange("O-S");
            const profitCols = parseColumnRange("T-X");
            const growthCols = parseColumnRange("Y-AE,AF-AL");
            const marketValueCol = parseColumnRange("C-D");
            const mainFundCols = parseColumnRange("L-N");
            const dailyChangeCol = parseColumnRange("E");
            const rangeChangeCols = parseColumnRange("F-K");
            const defaultFormatCols = parseColumnRange("E-AL");

            // 优化格式设置：减少循环次数和属性访问
            const numberFormats = new Map([
                [defaultFormatCols, "#,##0.00"],
                [[...revenueCols, ...profitCols, ...marketValueCol, ...mainFundCols], '0!.00,,"亿";-0!.00,,"亿"']
            ]);

            // 应用数字格式
            numberFormats.forEach((format, cols) => {
                cols.forEach(col => {
                    const column = targetSheet.Columns(col);
                    column.NumberFormat = format;
                    if (defaultFormatCols.includes(col)) {
                        column.ColumnWidth = 7;
                    }
                });
            });

            // 条件格式规则配置（集中管理规则）
            const conditionalFormats = [
                {
                    cols: rangeChangeCols,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 20, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 20, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_RED_FILL }
                    ]
                },
                {
                    cols: dailyChangeCol,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 8, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 8, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                },
                {
                    cols: growthCols,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 30, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 30, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_RED_FILL }
                    ]
                },
                {
                    cols: revenueCols,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 1000000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 1000000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                },
                {
                    cols: profitCols,
                    rules: [
                        { type: "less", value: 0, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_GREEN_FILL },
                        { type: "between", min: 0, max: 500000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 500000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                },
                {
                    cols: mainFundCols,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 50000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 50000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                }
            ];

            // 应用条件格式
            conditionalFormats.forEach(({ cols, rules }) => {
                cols.forEach(col => {
                    setConditionalFormatting(targetSheet, col, 3, lastRow, rules);
                });
            });

            // 手动计算一次
            app.Calculate();

            console.log("格式设置完成...");
            alert(`已自动用公式生成了"${targetSheetName}"，页面也设置了条件格式，还差最后一步，手动复制第一行的公式到后面行！`);
        } finally {
            // 恢复应用设置（确保无论是否出错都能恢复）
            app.Calculation = originalCalculation;
            app.ScreenUpdating = true; // 强制开启，确保界面刷新
            app.EnableEvents = originalEnableEvents;
        }
    } catch (e) {
        alert("操作失败：" + e.message);
        console.error(e);
        // 出错时确保恢复基本设置
        try {
            Application.ScreenUpdating = true;
            Application.Calculation = -4105; // 恢复自动计算
        } catch (err) { /* 忽略恢复错误 */ }
    }
}

function createMarketIndustrialTable() {
    try {
        // 获取当前工作簿
        const wb = Application.Workbooks.Item(1);
        if (!wb) {
            alert("请确保有打开的工作簿");
            return;
        }

        // 定义表名常量
        const targetSheetName = "行业板块";
        const sourceSheetName = "选股结果";

        // 检查原始数据表是否存在
        let sourceSheet;
        try {
            sourceSheet = wb.Worksheets.Item(sourceSheetName);
        } catch (e) {
            alert(`未找到"${sourceSheetName}"工作表，请先创建该表`);
            return;
        }

        // 检查概念板块是否存在，不存在则创建
        let targetSheet;
        try {
            targetSheet = wb.Worksheets.Item(targetSheetName);
        } catch (e) {
            alert(`未找到"${targetSheetName}"工作表，请先创建该表`);
            targetSheet = wb.Worksheets.Add();
            targetSheet.Name = targetSheetName;
        }

        // 性能优化设置
        const originalCalculation = wb.Application.Calculation;
        const originalScreenUpdating = wb.Application.ScreenUpdating;
        wb.Application.Calculation = 1; // 手动计算
        wb.Application.ScreenUpdating = false;

        try {
            // 查找行业列位置
            const industryColIndex = findColumnByHeader(sourceSheet, "所属同花顺行业");
            if (industryColIndex === -1) {
                debugFirstRowValues(sourceSheet);
                alert("未找到'所属同花顺行业'列");
                return false;
            }
            const industryColLetter = getColumnLetter(industryColIndex);

            // 获取有效数据范围（排除标题行）
            const sourceLastRow = sourceSheet.UsedRange.Rows.Count;
            const dataStartRow = 2; // 数据从第2行开始
            const dataEndRow = sourceLastRow;
            const biRange = `'${sourceSheetName}'!${industryColLetter}$${dataStartRow}:${industryColLetter}$${dataEndRow}`;

            // 定义公式配置，集中管理所有公式
            const formulaConfigs = [
                { cell: "B3", formula: `=COUNTIF(${biRange},"*"&A3&"*")` },
                { cell: "C3", formula: `=SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!F$${dataStartRow}:F$${dataEndRow})` },
                { cell: "D3", formula: `=IF(B3=0,0,C3/B3)` }, // 避免除以零
                { cell: "E3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!C$${dataStartRow}:C$${dataEndRow})/B3)` },//当日涨幅
                { cell: "F3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!G$${dataStartRow}:G$${dataEndRow})/B3)` },//区间涨幅
                { cell: "G3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!H$${dataStartRow}:H$${dataEndRow})/B3)` },
                { cell: "H3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!I$${dataStartRow}:I$${dataEndRow})/B3)` },
                { cell: "I3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!J$${dataStartRow}:J$${dataEndRow})/B3)` },
                { cell: "J3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!K$${dataStartRow}:K$${dataEndRow})/B3)` },
                { cell: "K3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!L$${dataStartRow}:L$${dataEndRow})/B3)` },
                { cell: "L3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!BC$${dataStartRow}:BC$${dataEndRow})/B3)` },
                { cell: "M3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!BD$${dataStartRow}:BD$${dataEndRow})/B3)` },
                { cell: "N3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!BE$${dataStartRow}:BE$${dataEndRow})/B3)` },
                { cell: "O3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!N$${dataStartRow}:N$${dataEndRow})/B3)` },
                { cell: "P3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!O$${dataStartRow}:O$${dataEndRow})/B3)` },
                { cell: "Q3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!P$${dataStartRow}:P$${dataEndRow})/B3)` },
                { cell: "R3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!Q$${dataStartRow}:Q$${dataEndRow})/B3)` },
                { cell: "S3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!R$${dataStartRow}:R$${dataEndRow})/B3)` },
                { cell: "T3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!S$${dataStartRow}:S$${dataEndRow})/B3)` },
                { cell: "U3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!T$${dataStartRow}:T$${dataEndRow})/B3)` },
                { cell: "V3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!U$${dataStartRow}:U$${dataEndRow})/B3)` },
                { cell: "W3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!V$${dataStartRow}:V$${dataEndRow})/B3)` },
                { cell: "X3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!W$${dataStartRow}:W$${dataEndRow})/B3)` },
                { cell: "Y3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!X$${dataStartRow}:X$${dataEndRow})/B3)` },
                { cell: "Z3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!Y$${dataStartRow}:Y$${dataEndRow})/B3)` },
                { cell: "AA3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!Z$${dataStartRow}:Z$${dataEndRow})/B3)` },
                { cell: "AB3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AA$${dataStartRow}:AA$${dataEndRow})/B3)` },
                { cell: "AC3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AB$${dataStartRow}:AB$${dataEndRow})/B3)` },
                { cell: "AD3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AC$${dataStartRow}:AC$${dataEndRow})/B3)` },
                { cell: "AE3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AD$${dataStartRow}:AD$${dataEndRow})/B3)` },
                { cell: "AF3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AF$${dataStartRow}:AF$${dataEndRow})/B3)` },
                { cell: "AG3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AG$${dataStartRow}:AG$${dataEndRow})/B3)` },
                { cell: "AH3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AH$${dataStartRow}:AH$${dataEndRow})/B3)` },
                { cell: "AI3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AI$${dataStartRow}:AI$${dataEndRow})/B3)` },
                { cell: "AJ3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AJ$${dataStartRow}:AJ$${dataEndRow})/B3)` },
                { cell: "AK3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AK$${dataStartRow}:AK$${dataEndRow})/B3)` },
                { cell: "AL3", formula: `=IF(B3=0,0,SUMIF(${biRange},"*"&A3&"*",'${sourceSheetName}'!AL$${dataStartRow}:AL$${dataEndRow})/B3)` }
            ];

            // 批量设置公式
            formulaConfigs.forEach(config => {
                targetSheet.Range(config.cell).Formula = config.formula;
            });

            // 自动填充公式（假设行业列表在A列，从A3开始有数据）
            const lastIndustryRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(-4162).Row; // xlUp = -4162
            if (lastIndustryRow > 3) {
                formulaConfigs.forEach(config => {
                    const sourceCell = targetSheet.Range(config.cell);
                    const fillRange = targetSheet.Range(
                        config.cell,
                        config.cell.replace("3", lastIndustryRow.toString())
                    );
                    sourceCell.AutoFill(fillRange);
                });
            }

            // 手动计算一次
            wb.Application.Calculate();

            // 颜色常量定义（BGR顺序）
            const COLORS = {
                TITLE_BG: RGB_BGR(51, 119, 255),
                TITLE_TEXT: RGB_BGR(255, 255, 255),
                GREEN_FILL: RGB_BGR(198, 239, 206),
                GREEN_TEXT: RGB_BGR(74, 148, 74),
                DARK_GREEN_FILL: RGB_BGR(0, 128, 0),
                YELLOW_FILL: RGB_BGR(255, 235, 156),
                YELLOW_TEXT: RGB_BGR(185, 141, 48),
                RED_FILL: RGB_BGR(255, 199, 206),
                RED_TEXT: RGB_BGR(191, 71, 78),
                DARK_RED_FILL: RGB_BGR(235, 0, 0),
                WHITE_TEXT: RGB_BGR(255, 255, 255)
            };

            // 列范围配置
            const columnRanges = {
                revenue: parseColumnRange("O-S"),
                profit: parseColumnRange("T-X"),
                growth: parseColumnRange("Y-AE,AF-AL"),
                marketValue: parseColumnRange("C-D"),
                mainFund: parseColumnRange("L-N"),
                dailyChange: parseColumnRange("E"),
                rangeChange: parseColumnRange("F-K"),
                defaultFormat: parseColumnRange("E-AL")
            };

            // 应用默认格式
            columnRanges.defaultFormat.forEach(col => {
                targetSheet.Columns(col).NumberFormat = "#,##0.00";
                targetSheet.Columns(col).ColumnWidth = 7;
            });

            // 应用亿单位格式
            [...columnRanges.revenue, ...columnRanges.profit, ...columnRanges.marketValue, ...columnRanges.mainFund]
                .forEach(col => {
                    targetSheet.Columns(col).NumberFormat = '0!.00,,"亿";-0!.00,,"亿"';
                });

            // 条件格式配置
            const conditionalFormats = [
                {
                    columns: columnRanges.rangeChange,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 20, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 20, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_RED_FILL }
                    ]
                },
                {
                    columns: columnRanges.dailyChange,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 8, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 8, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                },
                {
                    columns: columnRanges.growth,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 30, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 30, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_RED_FILL }
                    ]
                },
                {
                    columns: columnRanges.revenue,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 1000000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 1000000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                },
                {
                    columns: columnRanges.profit,
                    rules: [
                        { type: "less", value: 0, font: COLORS.WHITE_TEXT, fill: COLORS.DARK_GREEN_FILL },
                        { type: "between", min: 0, max: 500000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 500000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                },
                {
                    columns: columnRanges.mainFund,
                    rules: [
                        { type: "less", value: 0, font: COLORS.GREEN_TEXT, fill: COLORS.GREEN_FILL },
                        { type: "between", min: 0, max: 50000000, font: COLORS.YELLOW_TEXT, fill: COLORS.YELLOW_FILL },
                        { type: "greaterEqual", value: 50000000, font: COLORS.RED_TEXT, fill: COLORS.RED_FILL }
                    ]
                }
            ];

            // 应用条件格式
            conditionalFormats.forEach(item => {
                item.columns.forEach(col => {
                    setConditionalFormatting(targetSheet, col, 3, lastIndustryRow, item.rules);
                });
            });

            alert(`已自动生成"${targetSheetName}"并完成格式设置，公式已自动填充至所有行业行！`);
        } finally {
            // 恢复应用设置
            wb.Application.Calculation = originalCalculation;
            wb.Application.ScreenUpdating = originalScreenUpdating;
        }
    } catch (e) {
        alert("操作失败：" + e.message);
        console.error(e);
    }
}

/**
 * 自动过滤选结果自动过滤功能
 * 从"过滤条件"工作表读取条件，在"选股结果"工作表中应用筛选
 */
function autoFilterStockSelection() {
    try {
        // 获取当前工作簿
        const wb = Application.Workbooks.Item(1);
        if (!wb) {
            alert("请确保有打开的工作簿");
            return;
        }

        // 获取所需工作表
        const wsResult = wb.Worksheets.Item("选股结果");
        const wsCriteria = wb.Worksheets.Item("工作表首页");

        if (!wsResult || !wsCriteria) {
            alert("请确保存在名为'选股结果'和'过滤条件'的工作表");
            return;
        }

        // 定义选股结果表中的列映射
        const columnsMap = {
            industry: "D",         // 行业名称列
            concept: "BH",         // 所属概念列（后续会自动定位）
            marketValue: "F",      // 市值列
            revenue: "N",          // 营业收入列
            profit: "S",           // 营业利润列
            revenueGrowth: "X",    // 营收增长比例列
            profitGrowth: "AE",    // 利润增长比例列
            dayIncreaseRate: "C"   // 当日涨幅列
        };

        // 查找所属概念列的位置（处理可能的列位置变化）
        const tmpConceptCol = findColumnByHeader(wsResult, "所属概念");
        if (tmpConceptCol === -1) {
            debugFirstRowValues(wsResult);
            alert("表格没有找到[所属概念]列");
            return false;
        }
        console.log(`更新【所属概念】列为：${getColumnLetter(tmpConceptCol)}`);
        columnsMap.concept = getColumnLetter(tmpConceptCol);

        // 清除现有筛选状态
        if (wsResult.AutoFilterMode) {
            wsResult.AutoFilterMode = false;
        }
        wsResult.UsedRange.EntireRow.Hidden = false;

        // 从过滤条件表读取条件（统一处理空值）
        const criteria = {
            industry: getCellValue(wsCriteria.Range("B2")) || "",
            concept: getCellValue(wsCriteria.Range("B3")) || "",
            concept2: getCellValue(wsCriteria.Range("C3")) || "",
            concept3: getCellValue(wsCriteria.Range("D3")) || "",
            marketValueMin: getCellValue(wsCriteria.Range("B4")) || null,
            marketValueMax: getCellValue(wsCriteria.Range("B5")) || null,
            revenueValueMin: getCellValue(wsCriteria.Range("B6")) || null,
            profitValueMin: getCellValue(wsCriteria.Range("B7")) || null,
            revenueGrowthMin: getCellValue(wsCriteria.Range("B8")) || null,
            profitGrowthMin: getCellValue(wsCriteria.Range("B9")) || null,
            dayIncreaseRate: getCellValue(wsCriteria.Range("B10")) || null
        };

        // 设置筛选范围（使用第2行作为表头）
        const startCol = 1;
        const endCol = wsResult.UsedRange.Columns.Count;
        const headerRow = 2;
        const filterRange = wsResult.Range(
            wsResult.Cells(headerRow, startCol),
            wsResult.Cells(headerRow, endCol)
        );
        filterRange.AutoFilter();

        // 转换列字母为索引（统一处理列索引转换）
        const colIndexes = {
            industry: getColumnIndex(columnsMap.industry),
            concept: getColumnIndex(columnsMap.concept),
            marketValue: getColumnIndex(columnsMap.marketValue),
            revenue: getColumnIndex(columnsMap.revenue),
            profit: getColumnIndex(columnsMap.profit),
            revenueGrowth: getColumnIndex(columnsMap.revenueGrowth),
            profitGrowth: getColumnIndex(columnsMap.profitGrowth),
            dayIncreaseRate: getColumnIndex(columnsMap.dayIncreaseRate)
        };

        // 应用行业名称过滤
        if (criteria.industry) {
            filterRange.AutoFilter(colIndexes.industry, `*${criteria.industry}*`);
        }

        // 应用市值范围过滤
        if (criteria.marketValueMin !== null || criteria.marketValueMax !== null) {
            const minVal = parseFloat(criteria.marketValueMin);
            const maxVal = parseFloat(criteria.marketValueMax);
            const hasMin = !isNaN(minVal);
            const hasMax = !isNaN(maxVal);

            if (hasMin && hasMax) {
                filterRange.AutoFilter(colIndexes.marketValue, `>=${minVal}`, 1, `<=${maxVal}`);
            } else if (hasMin) {
                filterRange.AutoFilter(colIndexes.marketValue, `>=${minVal}`);
            } else if (hasMax) {
                filterRange.AutoFilter(colIndexes.marketValue, `<=${maxVal}`);
            }
        }

        // 通用数值过滤函数（减少重复代码）
        const applyNumberFilter = (criteriaValue, colIndex, operator = ">=") => {
            if (criteriaValue === null) return;
            const value = parseFloat(criteriaValue);
            if (!isNaN(value)) {
                filterRange.AutoFilter(colIndex, `${operator}${value}`);
            }
        };

        // 应用各数值列过滤
        applyNumberFilter(criteria.revenueValueMin, colIndexes.revenue);
        applyNumberFilter(criteria.dayIncreaseRate, colIndexes.dayIncreaseRate);
        applyNumberFilter(criteria.profitValueMin, colIndexes.profit);
        applyNumberFilter(criteria.revenueGrowthMin, colIndexes.revenueGrowth);
        applyNumberFilter(criteria.profitGrowthMin, colIndexes.profitGrowth);

        // 处理概念过滤
        const conceptConditions = [criteria.concept, criteria.concept2, criteria.concept3]
            .filter(cond => cond); // 过滤空条件

        if (conceptConditions.length > 0) {
            const targetColIndex = colIndexes.concept;
            const dataStartRow = 3;
            const dataLastRow = wsResult.UsedRange.Rows.Count;
            let matchCount = 0;

            console.log(`开始匹配概念条件：[${conceptConditions.join('], [')}]`);

            if (conceptConditions.length === 1) {
                // 单个条件使用AutoFilter效率更高
                filterRange.AutoFilter(targetColIndex, `*${conceptConditions[0]}*`);
            } else {
                // 多个条件使用行隐藏方式
                for (let row = dataStartRow; row <= dataLastRow; row++) {
                    try {
                        const cellValue = (getCellValue(wsResult.Cells(row, targetColIndex)) || "").toUpperCase();
                        const allMatch = conceptConditions.every(cond =>
                            cellValue.includes(cond.toUpperCase())
                        );

                        if (!allMatch) {
                            wsResult.Rows(row).Hidden = true;
                        } else {
                            matchCount++;
                        }
                    } catch (e) {
                        console.error(`处理行 ${row} 时出错: ${e.message}`);
                    }
                }
            }

            console.log(`满足概念组合的记录数：${matchCount}条`);
        }

        alert(`过滤完成！共应用了${countAppliedFilters(criteria)}个有效条件`);
    } catch (e) {
        alert(`执行过程中发生错误：${e.message}`);
        console.error(e);
    }
}



// 辅助函数：应用列格式
function applyColumnFormat(columns, formatOptions, sheet) {
    if (!columns || columns.length === 0) return;

    columns.forEach(col => {
        const column = sheet.Columns(col);
        if (formatOptions.numberFormat) {
            column.NumberFormat = formatOptions.numberFormat;
        }
        if (formatOptions.columnWidth !== undefined) {
            column.ColumnWidth = formatOptions.columnWidth;
        }
    });
}

// 辅助函数：批量应用条件格式
function applyConditionalFormattingToColumns(columns, sheet, startRow, lastRow, rules) {
    if (!columns || columns.length === 0) return;

    columns.forEach(col => {
        setConditionalFormatting(sheet, col, startRow, lastRow, rules);
    });
}