package com.amazing.poi.excel.read;

import com.amazing.poi.excel.cell.StringCellValueHandle;
import com.amazing.poi.excel.core.AbstractReadExcel;
import com.amazing.poi.excel.core.CellValueHandle;
import com.amazing.poi.excel.core.ValidateRow;
import com.amazing.poi.excel.exception.ErrorHeadException;
import com.amazing.poi.excel.exception.HandDataException;
import com.amazing.poi.excel.validata.NotValidateRow;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;


/**
 * <p>单个sheet的读取</p>
 *
 * @author amzing.
 * @version 0.1
 * @time 2019/10/22 23:05
 * @since 0.1
 */
public class SingleSheetRead extends AbstractReadExcel {

    private Logger logger = LoggerFactory.getLogger(SingleSheetRead.class);

    private List<Map<String, Object>> allData;
    private List<Map<String, Object>> correctData;
    private List<Map<String, Object>> errorData;

    private Map<String, String> headMap;

    /**
     * 是否过滤表头
     */
    private Boolean filterHead;

    private ValidateRow validate;

    private CellValueHandle cellValueHandle;

    private int headRows;

    private long cellSize;

    private StringBuilder errorMsg;

    public SingleSheetRead(InputStream inp, Map<String, String> headMap,
                           Boolean filterHead,
                           ValidateRow validate,
                           CellValueHandle cellValueHandle,
                           int headRows) throws IOException {
        super(inp);
        this.allData = new ArrayList<>();
        this.correctData = new ArrayList<>();
        this.errorData = new ArrayList<>();
        this.headMap = headMap;
        this.filterHead = filterHead;
        this.validate = validate;
        this.cellValueHandle = cellValueHandle;
        this.headRows = headRows;
        this.cellSize = headMap.size();
        this.errorMsg = new StringBuilder();
    }

    public SingleSheetRead(InputStream inp, Map<String, String> headMap,
                           int headRows) throws IOException {
        super(inp);
        this.allData = new ArrayList<>();
        this.correctData = new ArrayList<>();
        this.errorData = new ArrayList<>();
        this.headMap = headMap;
        this.filterHead = true;
        this.validate = new NotValidateRow();
        this.cellValueHandle = new StringCellValueHandle();
        this.headRows = headRows;
        this.cellSize = headMap.size();
        this.errorMsg = new StringBuilder();
    }

    public SingleSheetRead(InputStream inp) throws IOException {
        super(inp);
    }

    public void readSheet() throws HandDataException {
        Sheet sheet = getSheet(0);
        //获取表头所在的行
        Row row = sheet.getRow(headRows - 1);
        Iterator iterator = headMap.entrySet().iterator();

        for (long i = 0; i < cellSize; i++) {
            Cell tmp = row.getCell((int) i);
            String value = (String) cellValueHandle.getCellValue(tmp);

            Map.Entry m = (Map.Entry) iterator.next();
            if (!m.getValue().equals(value)) {
                throw new ErrorHeadException("error head map");
            }
        }

        super.readSheet(0);
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
            Object value = cellValueHandle.getCellValue(tmp);

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

        Object d = handleData(dataMap);
        handleAllData(d);

        //数据的正确性在数据转换后再做校验
        String validateMsg = (String) validate.validate(dataMap, rowNumber);
        if (StringUtils.isNotBlank(validateMsg)) {
            errorMsg.append(validateMsg);
            handleErrorData(d);
            return;
        }

        handleCorrectData(d);
    }

    public Object handleData(Object tempData) throws HandDataException {
        return tempData;
    }

    public void handleAllData(Object tempData) throws HandDataException {
        this.allData.add((Map<String, Object>) tempData);
    }

    public void handleCorrectData(Object tempData) throws HandDataException {
        this.correctData.add((Map<String, Object>) tempData);
    }

    public void handleErrorData(Object tempData) throws HandDataException {
        this.errorData.add((Map<String, Object>) tempData);
    }

    public List<Map<String, Object>> getAllData() {
        return allData;
    }

    public List<Map<String, Object>> getCorrectData() {
        return correctData;
    }

    public List<Map<String, Object>> getErrorData() {
        return errorData;
    }

    public String getErrorMsg() {
        return errorMsg.toString();
    }
}
