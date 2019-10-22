package com.amazing.poi.excel.core;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 11:14
 */
public abstract class AbstractReadExcel {

    private Logger logger = LoggerFactory.getLogger(AbstractReadExcel.class);

    private Workbook wb;

    public AbstractReadExcel(InputStream inp) throws IOException {
        wb = WorkbookFactory.create(inp);
    }

    public Workbook getWb() {
        return wb;
    }

    public void readSheet(int sheetIndex) {

        Sheet sheet = getSheet(sheetIndex);
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

        for (int i = 0; i <= lastRow; i++) {
            logger.debug("read ro number in {}", i);
            Row r = sheet.getRow(i);
            readRow(r,i);
        }
    }

    public Sheet getSheet(int index){
        return getWb().getSheetAt(index);
    }

    protected abstract void readRow(Row row, int rowNumber);

}
