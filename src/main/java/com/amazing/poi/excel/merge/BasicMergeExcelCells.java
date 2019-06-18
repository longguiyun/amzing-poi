package com.amazing.poi.excel.merge;

import com.amazing.poi.excel.core.MergeExcelCells;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Set;

/**
 * <p>
 * 简单的合并单元格
 * </p>
 *
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public class BasicMergeExcelCells implements MergeExcelCells {

    private final Logger LOGGER = LoggerFactory.getLogger(BasicMergeExcelCells.class);

    /**
     * 合并列的唯一标识，如果存在多个列合并，当标识已经结束合并，其它列也相应的结束合并
     */
    private String uniqueMergeCellIndex;

    /**
     * 需要合并的列
     */
    private Set<String> mergeCells;


    /**
     * 判断合并标识的值，如果不一致表示要合并数据
     */
    private Object tempData;

    private Integer tempBeginRow;

    /**
     * 合并开始行
     */
    private Integer mergeBeginRow;

    /**
     * 合并结束行
     */
    private Integer mergeEndRow;

    public BasicMergeExcelCells(String uniqueMergeCellIndex, Set<String> mergeCells) {
        this.uniqueMergeCellIndex = uniqueMergeCellIndex;
        this.mergeCells = mergeCells;
        this.mergeBeginRow = 0;
        this.mergeEndRow = 0;
    }

    @Override
    public void mergeCells(Sheet sheet, String columnName, Object data, int rowIndex, int columnIndex) {

        if (uniqueMergeCellIndex.equals(columnName)) {

            if (null == tempData) {
                tempData = data;
                tempBeginRow = rowIndex;
                mergeBeginRow = rowIndex;
                mergeEndRow = rowIndex;
            } else {
                if (!tempData.toString().equals(data.toString())) {
                    //需要换行
                    mergeBeginRow = tempBeginRow;
                    mergeEndRow = rowIndex - 1;
                    tempBeginRow = rowIndex;
                    tempData = data;
                }
            }

            LOGGER.debug("值:{} 当前行:{} 合并开始行:{} 合并结束行:{} 临时行:{}",data.toString(), rowIndex, mergeBeginRow, mergeEndRow, tempBeginRow);
        }

        if (mergeCells.contains(columnName) && mergeEndRow - mergeBeginRow > 0) {
            //开始合并数据
            sheet.addMergedRegion(new CellRangeAddress(
                    mergeBeginRow - 1, //first row (0-based)
                    mergeEndRow - 1, //last row  (0-based)
                    columnIndex, //first column (0-based)
                    columnIndex  //last column  (0-based)
            ));
        }
    }
}
