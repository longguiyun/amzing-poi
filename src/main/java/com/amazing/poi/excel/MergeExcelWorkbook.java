package com.amazing.poi.excel;

import com.amazing.poi.excel.core.MergeExcelCells;
import org.apache.poi.ss.usermodel.Row;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.instrument.IllegalClassFormatException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public class MergeExcelWorkbook extends SimpleExcelWorkbook {

    private MergeExcelCells mergeExcelCells;

    private Integer dataSize;

    public MergeExcelWorkbook(LinkedHashMap<String, String> headMap, MergeExcelCells mergeExcelCells,Integer dataSize) {
        super(headMap);
        this.mergeExcelCells = mergeExcelCells;
        this.dataSize = dataSize;
    }

    public MergeExcelWorkbook(LinkedHashMap<String, String> headMap, Integer rowIndex, MergeExcelCells mergeExcelCells, Integer dataSize) {
        super(headMap, rowIndex);
        this.mergeExcelCells = mergeExcelCells;
        this.dataSize = dataSize;
    }


    @Override
    public void fillContent(Object data) throws Exception {
        if(data instanceof Map){
            throw new IllegalClassFormatException("not support map, please use pojo");
        }

        PropertyDescriptor[] pds = null;
        List<Object> contentData = (List) data;
        for (int i = 0; i < contentData.size(); i++) {
            Object content = contentData.get(i);
            Row row = getAbstractExcelCell().createRow(getRowIndex());
            rowIndexIncrease();

            if(null == pds){
                //通过反射获取数据
                BeanInfo beanInfo = Introspector.getBeanInfo(content.getClass());
                pds = beanInfo.getPropertyDescriptors();
            }

            if(null == pds){
                throw new NullPointerException("fill content reflect error, target class "+content.getClass());
            }

            int colIndex = 0;
            Iterator itHeadMap = getHeadMap().entrySet().iterator();
            while (itHeadMap.hasNext()) {
                Map.Entry<String, String> s = (Map.Entry<String, String>) itHeadMap.next();
                String headKey = s.getKey();

                for (int i1 = 0; i1 < pds.length; i1++) {
                    PropertyDescriptor p = pds[i1];
                    if(headKey.equals(p.getName())){
                        Object value = p.getReadMethod().invoke(content);
                        getAbstractExcelCell().createCell(row, colIndex, value);
                        mergeExcelCells.mergeCells(getSheet(),headKey,value,getRowIndex(),colIndex);
                        colIndex++;
                        i1 = pds.length;
                    }
                }
            }
        }

        //最后一行的时候无法判断是否要合并，所以增加一个额外的虚拟行。用来判断最后一行是否需要合并
        if(getRowIndex().intValue() == dataSize+1){
            int colIndex = 0;
            rowIndexIncrease();
            Iterator itHeadMap = getHeadMap().entrySet().iterator();
            while (itHeadMap.hasNext()) {
                Map.Entry<String, String> s = (Map.Entry<String, String>) itHeadMap.next();
                String headKey = s.getKey();
                mergeExcelCells.mergeCells(getSheet(),headKey,"",getRowIndex(),colIndex);
                colIndex++;
            }
        }

    }
}
