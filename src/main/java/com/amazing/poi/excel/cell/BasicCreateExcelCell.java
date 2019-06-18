package com.amazing.poi.excel.cell;

import com.amazing.poi.excel.core.AbstractExcelCell;
import org.apache.poi.ss.usermodel.*;

import java.util.Calendar;
import java.util.Date;

/**
 * <p>
 *     创建简单的cell
 * </p>
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public class BasicCreateExcelCell extends AbstractExcelCell {

    public BasicCreateExcelCell(Workbook wb, Sheet sheet) {
        super(wb, sheet);
    }

    @Override
    public Row createRow(int rowIndex) {
        return getSheet().createRow(rowIndex);
    }

    @Override
    public Cell createCell(Row row, int colIndex, Object value) {
        Cell cell = this.createCell(row,colIndex);
        this.setCellValueAndFormat(cell,value);
        return cell;
    }

    /**
     * Creates a cell and aligns it a certain way.
     *
     * <p>
     *     用法：
     *     createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
     *     createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
     *     createCell(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
     *     createCell(wb, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
     *     createCell(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
     *     createCell(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
     *     createCell(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);
     *
     * </p>
     * @param row    the row to create the cell in
     * @param column the column number to create the cell in
     */
    private Cell createCell(Row row, int column) {
        Cell cell = row.createCell(column);
        CellStyle cellStyle = createSimpleCellStyle();
        cell.setCellStyle(cellStyle);
        return cell;
    }

    private CellStyle createSimpleCellStyle() {
        CellStyle cellStyle = super.getWb().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    private void setCellValue(Cell cell, Object value) {
        if (value.getClass() == Integer.class) {
            cell.setCellValue(Integer.parseInt(value.toString()));
        } else if (value.getClass() == Float.class) {
            cell.setCellValue(Float.parseFloat(value.toString()));
        } else if (value.getClass() == Long.class) {
            cell.setCellValue(Long.parseLong(value.toString()));
        } else if (value.getClass() == Double.class) {
            cell.setCellValue(Double.parseDouble(value.toString()));
        } else if (value.getClass() == String.class) {
            cell.setCellValue(value.toString());
        } else if (value.getClass() == Boolean.class) {
            cell.setCellValue(Boolean.parseBoolean(value.toString()));
        } else if (value.getClass() == Date.class) {
            cell.setCellValue((Date) value);
        } else if (value.getClass() == Calendar.class) {
            cell.setCellValue((Calendar) value);
        } else if (value.getClass() == RichTextString.class) {
            cell.setCellValue((RichTextString) value);
        } else {
            // nothing to do
        }
    }

    private void setCellValueAndFormat(Cell cell, Object value) {
        CreationHelper createHelper = super.getWb().getCreationHelper();
        CellStyle cellStyle = createSimpleCellStyle();

        if(null == value){
            // nothing to do
        } else if (value.getClass() == Integer.class) {
            cell.setCellValue(Integer.parseInt(value.toString()));
        } else if (value.getClass() == Float.class) {
            cell.setCellValue(Float.parseFloat(value.toString()));
        } else if (value.getClass() == Long.class) {
            cell.setCellValue(Long.parseLong(value.toString()));
        } else if (value.getClass() == Double.class) {
            cell.setCellValue(Double.parseDouble(value.toString()));
        } else if (value.getClass() == String.class) {
            cell.setCellValue(value.toString());
        } else if (value.getClass() == Boolean.class) {
            cell.setCellValue(Boolean.parseBoolean(value.toString()));
        } else if (value.getClass() == Date.class) {
            cellStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat("yyyy/m/d h:mm"));
            cell.setCellValue((Date) value);
            cell.setCellStyle(cellStyle);
        } else if (value.getClass() == Calendar.class) {
            cell.setCellValue((Calendar) value);
        } else if (value.getClass() == RichTextString.class) {
            cell.setCellValue((RichTextString) value);
        } else {
            // nothing to do
        }
    }

}
