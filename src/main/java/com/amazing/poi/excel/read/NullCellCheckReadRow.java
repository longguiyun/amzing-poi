package com.amazing.poi.excel.read;

import com.amazing.poi.excel.core.ValidateRow;

import java.util.Iterator;
import java.util.Map;

/**
 * Map键值对的行数据
 *
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 14:55
 */
public class NullCellCheckReadRow implements ValidateRow<Map<String, Object>> {

    private StringBuilder errorMsg;

    public NullCellCheckReadRow() {
        this.errorMsg = new StringBuilder();
    }

    @Override
    public void validate(Map<String, Object> map, int rowNumber) {
        Iterator iterator = map.entrySet().iterator();
        int cellNumber = 1;
        errorMsg.append("\n\t");
        while (iterator.hasNext()) {
            Map.Entry entry = (Map.Entry) iterator.next();
            if (null == entry.getValue()) {
                errorMsg.append(String.format("第%d行，第%d列,数据为空.", rowNumber, cellNumber));
            }

            cellNumber++;
        }
    }

    @Override
    public String errorMsg() {
        return this.errorMsg.toString();
    }
}
