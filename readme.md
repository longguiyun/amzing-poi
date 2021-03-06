# Apache POI工具类

### com.amazing.poi.excel
* 简介:
通过Apache POI实现了datagrid类型数据的excel表格导出

* 使用方法
    * 表头导出，使用LinkHashMap能保证表头的导出顺序
    ````
        String path = "url+文件名.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");
    
        SimpleExcelWorkbook sw = new SimpleExcelWorkbook(headMap);
        sw.fillHead();
        sw.write(path);
    ````
    
    * 表格数据导出
    ````
        String path = "url+文件名.xls";
        LinkedHashMap<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");
        //pojo类的字段必须要headMap中的字段对应，否者只会导出对应上headMap中的字段的数据
        //不支持Map类型的数据，因为字段为数字类型的时候,因为map可以存放Null数据
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
        //因为数据量大的时候，需要类似分页的情况来分批次处理数据
        sw.fillContent(contentData);
        sw.fillContent(contentData2);
        //表格宽度自适应
        sw.autoSizeColumn();
        sw.write(path);
    ````
    * 表格合并
    ````
        String path = "url+文件名.xls";
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
        //unique：合并的标识（所有的合并都依赖这个标识），mergeCell需要合并的字段
        BasicMergeExcelCells basicMergeExcelCells = new BasicMergeExcelCells(unique,mergeCell);

        MergeExcelWorkbook sw = new MergeExcelWorkbook(headMap,basicMergeExcelCells,contentData.size());
        sw.fillHead();
        sw.fillContent(contentData);
        sw.autoSizeColumn();
        sw.write(path);
    ````
    
    * 表格读取(返回List &lt;Map &gt;)
    ````
        //表头
        Map<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        String path = "C:\\Users\\amazing\\Desktop\\demo.xls";
        FileInputStream fis = new FileInputStream(path);
        //数据校验可以实现ValidateRow
        NullCellCheckReadRow readRowMap = new NullCellCheckReadRow();
        //数据转换可以实现CellValueHandle
        CellValueHandle cellValueHandle = new StringCellValueHandle();
        SingleSheetRead singleSheetRead = new SingleSheetRead(fis,headMap,true,readRowMap,cellValueHandle,1);

        singleSheetRead.readSheet();

        List<Map<String,Object>> data= singleSheetRead.getAllData();
        List<Map<String,Object>> correctData = singleSheetRead.getCorrectData();
        List<Map<String,Object>> errorData= singleSheetRead.getErrorData();

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

        System.out.println(singleSheetRead.getErrorMsg());
    ````
    
    * 表格读取(返回List &lt;Java Bean&gt;)
    ````
        //表头
        Map<String,String> headMap = new LinkedHashMap<>();
        headMap.put("name","姓名");
        headMap.put("age","年龄");
        headMap.put("gender","性别");

        String path = "C:\\Users\\amazing\\Desktop\\demo.xls";
        FileInputStream fis = new FileInputStream(path);

        //数据校验可以实现ValidateRow
        NullCellCheckReadRow readRowMap = new NullCellCheckReadRow();
        //数据转换可以实现CellValueHandle
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
    ````
