package com.amazing.poi.excel.commons;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Objects;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 17:52
 */
public class CellValueType {

    public static Object getCellValue(Cell cell) {
        Object value = null;

        if (Objects.isNull(cell)) {
            return null;
        }

        switch (cell.getCellType()) {
            case _NONE:
                break;
            case BLANK:
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case NUMERIC:
                value = cell.getNumericCellValue();
                break;
            case ERROR:
                value = cell.getErrorCellValue();
                break;
            default:

                break;
        }

        return value;
    }
}
