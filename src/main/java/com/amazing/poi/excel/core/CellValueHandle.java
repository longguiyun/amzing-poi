package com.amazing.poi.excel.core;

import org.apache.poi.ss.usermodel.Cell;

/**
 * @author amzing.
 * @version 0.1
 * @time 2019/10/22 23:31
 * @since 0.1
 */
public interface CellValueHandle {

    /**
     * 处理Cell的值
     *
     * @param cell
     * @return
     */
    Object getCellValue(Cell cell);
}
