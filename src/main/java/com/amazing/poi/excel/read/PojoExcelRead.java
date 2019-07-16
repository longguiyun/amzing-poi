package com.amazing.poi.excel.read;

import com.amazing.poi.excel.commons.ObjectCastUtils;
import com.amazing.poi.excel.core.ValidateRow;
import com.amazing.poi.excel.exception.HandDataException;

import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author amzing.
 * @version 0.1
 * @time 2019/7/15 22:47
 * @since 0.1
 */
public class PojoExcelRead<T extends Object> extends SimpleExcelRead {

    private List<T> data;

    private Class clazz;

    public PojoExcelRead(InputStream inp, Map head, Class clazz) {
        super(inp, head);
        this.data = new ArrayList<>();
        this.clazz = clazz;
    }

    public PojoExcelRead(InputStream inp, Map head, ValidateRow validate, Class clazz) {
        super(inp, head, validate);
        this.data = new ArrayList<>();
        this.clazz = clazz;
    }

    @Override
    public void handData(Map<String, Object> tempData) throws HandDataException {

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
            for (int i = 0; i < pds.length; i++) {
                PropertyDescriptor p = pds[i];
                String displayName = p.getDisplayName();
                Object value = tempData.get(displayName);

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

        data.add((T) o);
    }

    public List<T> getCastData() {
        return data;
    }
}
