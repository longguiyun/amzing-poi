package com.amazing.poi.excel.cell;

import com.amazing.poi.excel.core.CellValueHandle;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Objects;

/**
 * @author amzing.
 * @version 0.1
 * @time 2019/10/22 23:32
 * @since 0.1
 */
public class StringCellValueHandle implements CellValueHandle {


    @Override
    public Object getCellValue(Cell cell) {
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

        return Objects.isNull(value)? null : String.valueOf(value);
    }
}
