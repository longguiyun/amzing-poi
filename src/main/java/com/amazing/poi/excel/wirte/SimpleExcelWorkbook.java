package com.amazing.poi.excel.wirte;

import com.amazing.poi.excel.cell.BasicCreateExcelCell;
import com.amazing.poi.excel.core.AbstractExcelCell;
import com.amazing.poi.excel.core.AbstractExcelWorkbook;
import com.amazing.poi.excel.core.CreateExcelSheet;
import com.amazing.poi.excel.sheet.SingleExcelSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.instrument.IllegalClassFormatException;
import java.util.*;

/**
 * <p>
 *     简单的一行表头，一行数据，不存在合并单元格的操作
 * </p>
 *
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public class SimpleExcelWorkbook extends AbstractExcelWorkbook {

    /**
     * 按顺序存放表头数据，key：英文字段 value:中文别名
     */
    private LinkedHashMap<String, String> headMap;

    /**
     * 记录行数
     */
    private Integer rowIndex;

    private CreateExcelSheet createSheet;

    private Sheet sheet;

    private AbstractExcelCell abstractExcelCell;

    public SimpleExcelWorkbook(LinkedHashMap<String, String> headMap) {
        this.headMap = headMap;
        this.rowIndex = 0;
        this.createSheet = new SingleExcelSheet();
        this.sheet = this.createSheet.createSheet(getWb());
        this.abstractExcelCell = new BasicCreateExcelCell(getWb(), sheet);
    }

    public SimpleExcelWorkbook(LinkedHashMap<String, String> headMap, Integer rowIndex) {
        this.headMap = headMap;
        this.rowIndex = rowIndex;
        this.createSheet = new SingleExcelSheet();
        this.sheet = this.createSheet.createSheet(getWb());
        this.abstractExcelCell = new BasicCreateExcelCell(getWb(), sheet);
    }

    @Override
    public void fillHead() {
        Row row = abstractExcelCell.createRow(rowIndex);
        int colIndex = 0;
        Iterator itHeadMap = headMap.entrySet().iterator();
        while (itHeadMap.hasNext()) {
            Map.Entry<String, String> s = (Map.Entry<String, String>) itHeadMap.next();
            String headTitle = s.getValue();
            abstractExcelCell.createCell(row, colIndex, headTitle);
            colIndex++;
        }

        rowIndex++;
    }

    @SuppressWarnings("unchecked")
    @Override
    public void fillContent(Object data) throws Exception {
        if(data instanceof Map){
            throw new IllegalClassFormatException("not support map, please use pojo");
        }

        PropertyDescriptor[] pds = null;
        List<Object> contentData = (List) data;
        for (int i = 0; i < contentData.size(); i++) {
            Object content = contentData.get(i);
            Row row = abstractExcelCell.createRow(rowIndex);
            rowIndex++;

            if(null == pds){
                //通过反射获取数据
                BeanInfo beanInfo = Introspector.getBeanInfo(content.getClass());
                pds = beanInfo.getPropertyDescriptors();
            }

            if(null == pds){
                throw new NullPointerException("fill content reflect error, target class "+content.getClass());
            }

            int colIndex = 0;
            Iterator itHeadMap = headMap.entrySet().iterator();
            while (itHeadMap.hasNext()) {
                Map.Entry<String, String> s = (Map.Entry<String, String>) itHeadMap.next();
                String headKey = s.getKey();

                for (int i1 = 0; i1 < pds.length; i1++) {
                    PropertyDescriptor p = pds[i1];
                    if(headKey.equals(p.getName())){
                        Object value = p.getReadMethod().invoke(content);
                        abstractExcelCell.createCell(row, colIndex, value);

                        colIndex++;
                        i1 = pds.length;
                    }
                }
            }
        }
    }

    @Override
    public void autoSizeColumn() {
        int colIndex = 0;
        Iterator itHeadMap = headMap.entrySet().iterator();
        while (itHeadMap.hasNext()) {
            Map.Entry<String, String> s = (Map.Entry<String, String>) itHeadMap.next();
            String headTitle = s.getValue();
            //adjust column width to fit the content
            sheet.autoSizeColumn(colIndex);
            colIndex++;
        }
    }

    public LinkedHashMap<String, String> getHeadMap() {
        return headMap;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public void rowIndexIncrease(){
        this.rowIndex++;
    }

    public CreateExcelSheet getCreateSheet() {
        return createSheet;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public AbstractExcelCell getAbstractExcelCell() {
        return abstractExcelCell;
    }
}
