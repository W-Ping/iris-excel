package com.iris.excelfile.utils;

import com.alibaba.fastjson.JSON;
import com.iris.excelfile.exception.ExcelParseException;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 18:46
 * @see
 */
public class JSONUtil {
    public static String objectToString(Object object) {
        return JSON.toJSONString(object);
    }

    /**
     * @param source
     * @param cls
     * @param <V>
     * @return
     */
    public static <V, T> List<T> convertMapToList(List<Map<String, V>> list, Class<T> cls) {
        List<T> resultList = new ArrayList<>();
        if (list != null && list.size() > 0 && cls != null) {
            for (Map<String, V> source : list) {
                T t = convertMapToObject(source, cls);
                if (t != null) {
                    resultList.add(t);
                }
            }
        }
        return resultList;
    }

    public static <V, T> T convertMapToObject(Map<String, V> source, Class<T> cls) {
        if (source == null || cls == null) {
            return null;
        }
        try {
            Map<String, Field> objectAllFieldMap = FieldUtil.getObjectAllFieldMap(cls);
            T instance = cls.newInstance();
            for (Map.Entry<String, V> map : source.entrySet()) {
                String key = toCamelCase(map.getKey(), true);
                V value = map.getValue();
                Field field = objectAllFieldMap.get(key);
                FieldUtil.setValueToField(instance, field, value);
            }
            return instance;
        } catch (Exception e) {
            throw new ExcelParseException("map cast to " + cls.getSimpleName() + " fail!");
        }
    }

    public static String toCamelCase(String name, boolean flag) {
        if (name.indexOf("_") != -1 && flag) {
            String[] strArr = name.split("_");
            StringBuilder sb = new StringBuilder();
            String s = null;
            for (int i = 0; i < strArr.length; i++) {
                s = strArr[i];
                if (i == 0) {
                    s = s.replaceFirst(s.substring(0, 1), s.substring(0, 1).toLowerCase());
                } else {
                    s = s.replaceFirst(s.substring(0, 1), s.substring(0, 1).toUpperCase());
                }
                sb.append(s);
            }
            name = sb.toString();
        }
        return name;
    }
}
