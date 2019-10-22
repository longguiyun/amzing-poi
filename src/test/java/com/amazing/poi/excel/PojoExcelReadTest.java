package com.amazing.poi.excel;

import com.amazing.poi.excel.cell.StringCellValueHandle;
import com.amazing.poi.excel.core.CellValueHandle;
import com.amazing.poi.excel.read.PojoExcelRead;
import com.amazing.poi.excel.validata.NullCellCheckReadRow;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author amzing.
 * @version 0.1
 * @time 2019/10/23 0:04
 * @since 0.1
 */
public class PojoExcelReadTest {

    @Test
    public void convertPojoTest() throws IOException {

        Map<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        String path = "C:\\Users\\amazing\\Desktop\\demo.xls";
        FileInputStream fis = new FileInputStream(path);

        NullCellCheckReadRow readRowMap = new NullCellCheckReadRow();
        CellValueHandle cellValueHandle = new StringCellValueHandle();

        PojoExcelRead<SimplePOJO> pojoPojoExcelRead = new PojoExcelRead<>(fis,headMap,
                true,readRowMap,cellValueHandle,1,SimplePOJO.class);

        pojoPojoExcelRead.readSheet();

        List<SimplePOJO> data= pojoPojoExcelRead.getCastAllData();
        List<SimplePOJO> correctData = pojoPojoExcelRead.getCastCorrectData();
        List<SimplePOJO> errorData= pojoPojoExcelRead.getCastErrorData();

        System.err.println("全部数据");
        for (int i = 0; i < data.size(); i++) {
            System.out.println(data.get(i).toString());
        }

        System.err.println("正确数据");
        for (int i = 0; i < correctData.size(); i++) {
            System.out.println(correctData.get(i).toString());
        }

        System.err.println("错误数据");
        for (int i = 0; i < errorData.size(); i++) {
            System.out.println(errorData.get(i).toString());
        }

        System.out.println(pojoPojoExcelRead.getErrorMsg());
    }
}
