package com.amazing.poi.excel.core;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p>
 *     创建sheet接口
 * </p>
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public interface CreateExcelSheet {

    Sheet createSheet(Workbook wb);
}
