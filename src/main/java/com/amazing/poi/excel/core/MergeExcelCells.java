package com.amazing.poi.excel.core;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * <p>
 *     合并单元格接口
 * </p>
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public interface MergeExcelCells {

    void mergeCells(Sheet sheet, String columnName, Object data, int rowIndex, int columnIndex);
}
