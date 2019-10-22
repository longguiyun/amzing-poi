package com.amazing.poi.excel.validata;

import com.amazing.poi.excel.core.ValidateRow;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 14:59
 */
public class NotValidateRow implements ValidateRow {

    @Override
    public Object validate(Object o, int rowNumber) {
        return null;
    }


}
