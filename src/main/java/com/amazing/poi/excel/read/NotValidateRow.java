package com.amazing.poi.excel.read;

import com.amazing.poi.excel.core.ValidateRow;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 14:59
 */
public class NotValidateRow implements ValidateRow {

    @Override
    public void validate(Object o, int rowNumber) {
        return;
    }

    @Override
    public String errorMsg() {
        return null;
    }
}
