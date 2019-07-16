package com.amazing.poi.excel.core;


/**
 *
 * @param <T> 解析后的行数据类型
 *
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 14:50
 */
public interface ValidateRow<T extends Object> {

    void validate(T t, int rowNumber);

    String errorMsg();
}
