package com.amazing.poi.excel.read;

import com.amazing.poi.excel.commons.ObjectCastUtils;
import com.amazing.poi.excel.core.CellValueHandle;
import com.amazing.poi.excel.core.ValidateRow;
import com.amazing.poi.excel.exception.HandDataException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * @author amzing.
 * @version 0.1
 * @time 2019/7/15 22:47
 * @since 0.1
 */
public class PojoExcelRead<T extends Object> extends SingleSheetRead {

    private Logger logger = LoggerFactory.getLogger(SingleSheetRead.class);

    private List<T> castAllData;
    private List<T> castCorrectData;
    private List<T> castErrorData;

    private Class clazz;

    public PojoExcelRead(InputStream inp,
                         Map<String, String> headMap,
                         Boolean filterHead,
                         ValidateRow validate,
                         CellValueHandle cellValueHandle,
                         int headRows,
                         Class clazz) throws IOException {
        super(inp, headMap, filterHead, validate, cellValueHandle, headRows);
        this.castAllData = new ArrayList<>();
        this.castCorrectData = new ArrayList<>();
        this.castErrorData = new ArrayList<>();
        this.clazz = clazz;
    }

    public PojoExcelRead(InputStream inp, Map<String, String> headMap, int headRows, Class clazz) throws IOException {
        super(inp, headMap, headRows);
        this.castAllData = new ArrayList<>();
        this.castCorrectData = new ArrayList<>();
        this.castErrorData = new ArrayList<>();
        this.clazz = clazz;
    }

    public PojoExcelRead(InputStream inp, Class clazz) throws IOException {
        super(inp);
        this.castAllData = new ArrayList<>();
        this.castCorrectData = new ArrayList<>();
        this.castCorrectData = new ArrayList<>();
        this.clazz = clazz;
    }

    @Override
    public Object handleData(Object tempData) throws HandDataException {
        Map<String,Object> mapData = (Map<String, Object>) tempData;
        Object o = null;

        PropertyDescriptor[] pds = null;

        try {
            Class c = Class.forName(clazz.getName());
            o = c.newInstance();
            //通过反射获取数据
            BeanInfo beanInfo = Introspector.getBeanInfo(o.getClass());
            pds = beanInfo.getPropertyDescriptors();

        } catch (IntrospectionException e) {
            throw new HandDataException("reflect get bean error ", e);
        } catch (ClassNotFoundException e) {
            throw new HandDataException("reflect get bean error ", e);
        } catch (IllegalAccessException e) {
            throw new HandDataException(e);
        } catch (InstantiationException e) {
            throw new HandDataException(e);
        }

        if (null == o) {
            throw new HandDataException("can not get bean info");
        }

        try {
            //通过反射构造java对象
            for (int i = 0; i < pds.length; i++) {
                PropertyDescriptor p = pds[i];
                String displayName = p.getDisplayName();
                Object value = mapData.get(displayName);

                if (null == value) {
                    continue;
                }
                Class type = p.getPropertyType();
                Object castValue = ObjectCastUtils.objectCast(type,value);
                p.getWriteMethod().invoke(o, castValue);
            }
        } catch (IllegalAccessException e) {
            throw new HandDataException("hand data Illegal", e);
        } catch (InvocationTargetException e) {
            throw new HandDataException("hand data invocation error", e);
        }

        return (T) o;
    }

    @Override
    public void handleAllData(Object tempData) throws HandDataException {
        this.castAllData.add((T) tempData);
    }

    @Override
    public void handleCorrectData(Object tempData) throws HandDataException {
        this.castCorrectData.add((T) tempData);
    }

    @Override
    public void handleErrorData(Object tempData) throws HandDataException {
        this.castErrorData.add((T) tempData);
    }

    public List<T> getCastAllData() {
        return castAllData;
    }

    public List<T> getCastCorrectData() {
        return castCorrectData;
    }

    public List<T> getCastErrorData() {
        return castErrorData;
    }
}
