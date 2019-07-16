package com.amazing.poi.excel;

import com.amazing.poi.excel.read.NullCellCheckReadRow;
import com.amazing.poi.excel.read.PojoExcelRead;
import com.amazing.poi.excel.read.SimpleExcelRead;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/7/15 14:05
 */
public class ReadExcelTest {

    @Test
    public void simpleReadTest() throws IOException {

        Map<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        String path = "C:\\Users\\amazing\\Desktop\\demo.xls";
        FileInputStream fis = new FileInputStream(path);

        NullCellCheckReadRow readRowMap = new NullCellCheckReadRow();
        SimpleExcelRead simpleExcelRead = new SimpleExcelRead(fis,headMap,readRowMap);

        simpleExcelRead.readSingleSheet();

        List<Map<String,Object>> data= simpleExcelRead.getData();

        System.out.println();

        for (int i = 0; i < data.size(); i++) {
            System.out.println(data.get(i).toString());
        }

        System.out.println(readRowMap.errorMsg());
    }

    @Test
    public void convertPojoTest() throws IOException {

        Map<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        String path = "C:\\Users\\amazing\\Desktop\\demo.xls";
        FileInputStream fis = new FileInputStream(path);

        NullCellCheckReadRow readRowMap = new NullCellCheckReadRow();

        PojoExcelRead<SimplePOJO> pojoPojoExcelRead = new PojoExcelRead<>(fis,headMap,readRowMap,SimplePOJO.class);
        pojoPojoExcelRead.readSingleSheet();
        List<SimplePOJO> list = pojoPojoExcelRead.getCastData();

        System.out.println(readRowMap.errorMsg());
        list.stream().forEach(l->{
            System.out.println(l.toString());
        });
    }
}
