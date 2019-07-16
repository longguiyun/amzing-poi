package com.amazing.poi.excel.read;

import com.amazing.poi.excel.core.AbstractReadExcel;
import com.amazing.poi.excel.core.ValidateRow;
import com.amazing.poi.excel.exception.HandDataException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.*;

import static com.amazing.poi.excel.commons.CellValueType.getCellValue;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 11:42
 */
public class SimpleExcelRead extends AbstractReadExcel {

    private Logger logger = LoggerFactory.getLogger(SimpleExcelRead.class);

    private List<Map<String, Object>> data;

    private Map<String, String> headMap;

    /**
     * 是否过滤表头
     */
    private Boolean filterHead;

    private ValidateRow validate;

    private int headRows;

    private long cellSize;

    public SimpleExcelRead(InputStream inp, Map head) {
        super(inp);
        this.data = new LinkedList<>();
        this.headMap = head;
        this.filterHead = true;
        this.cellSize = head.size();
        this.headRows = 1;
        this.validate = new NotValidateRow();
    }

    public SimpleExcelRead(InputStream inp, Map head, ValidateRow validate) {
        super(inp);
        this.data = new LinkedList<>();
        this.headMap = head;
        this.filterHead = true;
        this.cellSize = head.size();
        this.headRows = 1;
        this.validate = validate;
    }

    public List<Map<String, Object>> getData() {
        return data;
    }

    @Override
    protected void readRow(Row row, int rowNumber) {

        if (Objects.isNull(row)) {
            logger.warn("row data is null");
            return;
        }

        if (filterHead && rowNumber < headRows) {
            return;
        }

        Map<String, Object> dataMap = new HashMap<>(1);

        Iterator iterator = null;
        for (long i = 0; i < cellSize; i++) {
            Cell tmp = row.getCell((int) i);
            Object value = getCellValue(tmp);

            if (i == 0) {
                iterator = headMap.entrySet().iterator();
            }

            if (Objects.isNull(iterator)) {
                throw new IllegalStateException("can not read head map");
            }

            Map.Entry entry = (Map.Entry) iterator.next();

            dataMap.put(entry.getKey().toString(), value);
            logger.debug("cell number {}, value is {}", i, value);
        }

        validate.validate(dataMap, rowNumber);

        handData(dataMap);
    }

    public void handData(Map<String,Object> tempData) throws HandDataException {
        this.data.add(tempData);
    }


}
