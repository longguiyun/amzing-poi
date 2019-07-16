package com.amazing.poi.excel.commons;

import javax.xml.crypto.Data;
import java.text.DecimalFormat;
import java.text.NumberFormat;

/**
 * @author amzing.
 * @version 0.1
 * @time 2019/7/16 23:04
 * @since 0.1
 */
public class ObjectCastUtils {

    public static Object objectCast(Class clz, Object o){

        Object target;
        if(clz == int.class || clz == Integer.class){
            DecimalFormat d = new DecimalFormat();
            target = Integer.valueOf(d.format(o));
        } else if(clz == short.class || clz == Short.class){
            DecimalFormat d = new DecimalFormat();
            target = Short.valueOf(d.format(o));
        } else if(clz == double.class || clz == Double.class){
            DecimalFormat d = new DecimalFormat();
            target = Double.valueOf(d.format(o));
        } else if(clz == long.class || clz == Long.class){
            DecimalFormat d = new DecimalFormat();
            target = Long.valueOf(d.format(o));
        } else if(clz == float.class || clz == Float.class){
            DecimalFormat d = new DecimalFormat();
            target = Float.valueOf(d.format(o));
        } else if(clz == String.class || clz == char.class){
            target = String.valueOf(o);
        } else {
            target = o;
        }

        return target;
    }
}
