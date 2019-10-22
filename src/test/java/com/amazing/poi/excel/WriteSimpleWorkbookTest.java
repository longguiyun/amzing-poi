package com.amazing.poi.excel;

import com.amazing.poi.excel.merge.BasicMergeExcelCells;
import com.amazing.poi.excel.wirte.MergeExcelWorkbook;
import com.amazing.poi.excel.wirte.SimpleExcelWorkbook;
import org.junit.Test;

import java.io.IOException;
import java.util.*;

/**
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public class WriteSimpleWorkbookTest {

    @Test
    public void fillHeadTest() throws IOException {

        String path = "C:\\Users\\amazing\\Desktop\\m-demo-fill-head.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        SimpleExcelWorkbook sw = new SimpleExcelWorkbook(headMap);
        sw.fillHead();
        sw.write(path);
    }

    @Test
    public void fillContentTest() throws Exception {

        String path = "C:\\Users\\amazing\\Desktop\\m-fill-content.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        List<SimplePOJO> contentData = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            SimplePOJO pojo = new SimplePOJO();
            pojo.setName("name:"+i);
            pojo.setGender("gender:"+i);
            contentData.add(pojo);
        }

        List<SimplePOJO> contentData2 = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            SimplePOJO pojo = new SimplePOJO();
            pojo.setName("name2:"+i);
            pojo.setAge(i);
            pojo.setGender("gender2:"+i);
            contentData2.add(pojo);
        }

        SimpleExcelWorkbook sw = new SimpleExcelWorkbook(headMap);
        sw.fillHead();
        sw.fillContent(contentData);
        sw.fillContent(contentData2);
        sw.autoSizeColumn();
        sw.write(path);
    }

    @Test
    public void mergeCellsTest() throws Exception {

        String path = "C:\\Users\\amazing\\Desktop\\m-cell-1.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        List<SimplePOJO> contentData = new ArrayList<>();

        SimplePOJO pojo1 = new SimplePOJO();
        pojo1.setName("name:1");
        pojo1.setAge(0);
        pojo1.setGender("gender:1");

        SimplePOJO pojo2 = new SimplePOJO();
        pojo2.setName("name:2");
        pojo2.setAge(0);
        pojo2.setGender("gender:2");

        SimplePOJO pojo3 = new SimplePOJO();
        pojo3.setName("name:3");
        pojo3.setAge(0);
        pojo3.setGender("gender:3");

        //一开始就合并
        contentData.add(pojo1);
        contentData.add(pojo2);
        contentData.add(pojo3);

        String unique = "name";
        Set<String> mergeCell = new HashSet<>();
        mergeCell.add("name");

        BasicMergeExcelCells basicMergeExcelCells = new BasicMergeExcelCells(unique,mergeCell);

        MergeExcelWorkbook sw = new MergeExcelWorkbook(headMap,basicMergeExcelCells,contentData.size());
        sw.fillHead();
        sw.fillContent(contentData);
        sw.autoSizeColumn();
        sw.write(path);
    }

    @Test
    public void mergeCellsTest2() throws Exception {

        String path = "C:\\Users\\amazing\\Desktop\\m-cell-2.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        List<SimplePOJO> contentData = new ArrayList<>();

        SimplePOJO pojo1 = new SimplePOJO();
        pojo1.setName("name:1");
        pojo1.setAge(0);
        pojo1.setGender("gender:1");

        SimplePOJO pojo2 = new SimplePOJO();
        pojo2.setName("name:1");
        pojo2.setAge(0);
        pojo2.setGender("gender:2");

        SimplePOJO pojo3 = new SimplePOJO();
        pojo3.setName("name:2");
        pojo3.setAge(0);
        pojo3.setGender("gender:3");

        //一开始就合并
        contentData.add(pojo1);
        contentData.add(pojo2);
        contentData.add(pojo3);

        String unique = "name";
        Set<String> mergeCell = new HashSet<>();
        mergeCell.add("name");

        BasicMergeExcelCells basicMergeExcelCells = new BasicMergeExcelCells(unique,mergeCell);

        MergeExcelWorkbook sw = new MergeExcelWorkbook(headMap,basicMergeExcelCells,contentData.size());
        sw.fillHead();
        sw.fillContent(contentData);
        //sw.autoSizeColumn();
        sw.write(path);
    }

    @Test
    public void mergeCellsTest3() throws Exception {

        String path = "C:\\Users\\amazing\\Desktop\\m-cell-3.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        List<SimplePOJO> contentData = new ArrayList<>();

        SimplePOJO pojo1 = new SimplePOJO();
        pojo1.setName("name:1");
        pojo1.setAge(0);
        pojo1.setGender("gender:1");

        SimplePOJO pojo2 = new SimplePOJO();
        pojo2.setName("name:2");
        pojo2.setAge(0);
        pojo2.setGender("gender:2");

        SimplePOJO pojo3 = new SimplePOJO();
        pojo3.setName("name:2");
        pojo3.setAge(0);
        pojo3.setGender("gender:3");

        SimplePOJO pojo4 = new SimplePOJO();
        pojo4.setName("name:4");
        pojo4.setAge(0);
        pojo4.setGender("gender:4");

        //一开始就合并
        contentData.add(pojo1);
        contentData.add(pojo2);
        contentData.add(pojo3);
        contentData.add(pojo4);

        String unique = "name";
        Set<String> mergeCell = new HashSet<>();
        mergeCell.add("name");

        BasicMergeExcelCells basicMergeExcelCells = new BasicMergeExcelCells(unique,mergeCell);

        MergeExcelWorkbook sw = new MergeExcelWorkbook(headMap,basicMergeExcelCells,contentData.size());
        sw.fillHead();
        sw.fillContent(contentData);
        //sw.autoSizeColumn();
        sw.write(path);
    }

    @Test
    public void mergeCellsTest4() throws Exception {

        String path = "C:\\Users\\amazing\\Desktop\\m-cell-4.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        List<SimplePOJO> contentData = new ArrayList<>();

        SimplePOJO pojo1 = new SimplePOJO();
        pojo1.setName("name:1");
        pojo1.setAge(0);
        pojo1.setGender("gender:1");

        SimplePOJO pojo2 = new SimplePOJO();
        pojo2.setName("name:2");
        pojo2.setAge(0);
        pojo2.setGender("gender:2");

        SimplePOJO pojo3 = new SimplePOJO();
        pojo3.setName("name:2");
        pojo3.setAge(0);
        pojo3.setGender("gender:3");

        //一开始就合并
        contentData.add(pojo1);
        contentData.add(pojo2);
        contentData.add(pojo3);

        String unique = "name";
        Set<String> mergeCell = new HashSet<>();
        mergeCell.add("name");

        BasicMergeExcelCells basicMergeExcelCells = new BasicMergeExcelCells(unique,mergeCell);

        MergeExcelWorkbook sw = new MergeExcelWorkbook(headMap,basicMergeExcelCells,contentData.size());
        sw.fillHead();
        sw.fillContent(contentData);
        //sw.autoSizeColumn();
        sw.write(path);
    }

    @Test
    public void mergeCellsTest5() throws Exception {

        String path = "C:\\Users\\amazing\\Desktop\\m-cell-5.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        List<SimplePOJO> contentData = new ArrayList<>();

        SimplePOJO pojo1 = new SimplePOJO();
        pojo1.setName("name:1");
        pojo1.setAge(0);
        pojo1.setGender("gender:1");

        SimplePOJO pojo2 = new SimplePOJO();
        pojo2.setName("name:1");
        pojo2.setAge(0);
        pojo2.setGender("gender:2");

        SimplePOJO pojo3 = new SimplePOJO();
        pojo3.setName("name:3");
        pojo3.setAge(0);
        pojo3.setGender("gender:3");

        SimplePOJO pojo4 = new SimplePOJO();
        pojo4.setName("name:4");
        pojo4.setAge(0);
        pojo4.setGender("gender:4");

        SimplePOJO pojo5 = new SimplePOJO();
        pojo5.setName("name:4");
        pojo5.setAge(0);
        pojo5.setGender("gender:5");

        SimplePOJO pojo6 = new SimplePOJO();
        pojo6.setName("name:6");
        pojo6.setAge(0);
        pojo6.setGender("gender:6");

        SimplePOJO pojo7 = new SimplePOJO();
        pojo7.setName("name:7");
        pojo7.setAge(0);
        pojo7.setGender("gender:7");

        SimplePOJO pojo8 = new SimplePOJO();
        pojo8.setName("name:7");
        pojo8.setAge(0);
        pojo8.setGender("gender:8");

        //一开始就合并
        contentData.add(pojo1);
        contentData.add(pojo2);
        contentData.add(pojo3);
        contentData.add(pojo4);
        contentData.add(pojo5);
        contentData.add(pojo6);
        contentData.add(pojo7);
        contentData.add(pojo8);
        contentData.add(pojo8);

        String unique = "name";
        Set<String> mergeCell = new HashSet<>();
        mergeCell.add("name");
        mergeCell.add("gender");

        BasicMergeExcelCells basicMergeExcelCells = new BasicMergeExcelCells(unique,mergeCell);

        MergeExcelWorkbook sw = new MergeExcelWorkbook(headMap,basicMergeExcelCells,contentData.size());
        sw.fillHead();
        sw.fillContent(contentData);
        //sw.autoSizeColumn();
        sw.write(path);
    }

}
