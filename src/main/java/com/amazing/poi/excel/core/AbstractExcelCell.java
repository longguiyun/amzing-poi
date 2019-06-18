package com.amazing.poi.excel.core;

import org.apache.poi.ss.usermodel.*;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public abstract class AbstractExcelCell {

    private Workbook wb;

    private Sheet sheet;

    public AbstractExcelCell(Workbook wb, Sheet sheet) {
        this.wb = wb;
        this.sheet = sheet;
    }

    public abstract Row createRow(int rowIndex);

    public abstract Cell createCell(Row row, int colIndex, Object value);

    public Workbook getWb() {
        return wb;
    }

    public Sheet getSheet() {
        return sheet;
    }
}
