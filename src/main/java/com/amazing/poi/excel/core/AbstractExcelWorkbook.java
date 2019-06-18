package com.amazing.poi.excel.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * <p>
 *      Excel POI基础类
 * </p>
 * @author longguiyun
 * @version 1.0
 * @date 2019/6/17 20:10
 */
public abstract class AbstractExcelWorkbook {

    private Workbook wb;

    public AbstractExcelWorkbook() {
        this.wb = new XSSFWorkbook();
    }

    public AbstractExcelWorkbook(Workbook wb) {
        this.wb = wb;
    }

    /**
     * <p>
     *     填充excel表头数据
     * </p>
     */
    public abstract void fillHead();

    /**
     * <p>
     *     填充excel表内内容数据
     * </p>
     * @param data
     * @throws Exception
     */
    public abstract void fillContent(Object data) throws Exception;

    /**
     * <p>
     *     自动调整excel的column宽度。默认不实现调整。需要调整的在子类中实现具体细节
     * </p>
     */
    public void autoSizeColumn(){
        //default not auto set width of column
    }

    public void write(String path) throws IOException {

        OutputStream fs = null;
        try {
            fs = new FileOutputStream(path);
            wb.write(fs);
            fs.flush();
        } catch (IOException e) {
            throw e;
        } finally {
            if (null != fs) {
                fs.close();
            }

            if (null != wb) {
                wb.close();
            }
        }
    }

    public void write(OutputStream os) throws IOException {
        try {
            wb.write(os);
            os.flush();
        } catch (IOException e) {
            throw e;
        } finally {
            if (null != os) {
                os.close();
            }

            if (null != wb) {
                wb.close();
            }
        }
    }

    public Workbook getWb() {
        return wb;
    }
}
