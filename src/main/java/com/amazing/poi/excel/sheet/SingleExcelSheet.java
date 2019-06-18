package com.amazing.poi.excel.sheet;

import com.amazing.poi.excel.core.CreateExcelSheet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;


/**
 * <p>
 *     创建单个sheet
 * </p>
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public class SingleExcelSheet implements CreateExcelSheet {

    @Override
    public Sheet createSheet(Workbook wb) {
        return this.create(wb,"sheet1");
    }

    public Sheet create(Workbook wb, String sheetName) {

        // Note that sheet name is Excel must not exceed 31 characters
        // and must not contain any of the any of the following characters:
        // 0x0000
        // 0x0003
        // colon (:)
        // backslash (\)
        // asterisk (*)
        // question mark (?)
        // forward slash (/)
        // opening square bracket ([)
        // closing square bracket (])

        // You can use org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String nameProposal)}
        // for a safe way to create valid names, this utility replaces invalid characters with a space (' ')
        // returns " O'Brien's sales   "
        String safeName = WorkbookUtil.createSafeSheetName(sheetName);
        return wb.createSheet(safeName);
    }
}
