package com.iris.excelfile.utils;

import com.iris.excelfile.annotation.ExcelWriteProperty;
import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.BaseColumnProperty;
import com.iris.excelfile.metadata.BaseRowModel;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 16:58
 * @see
 */
@Slf4j
public class FieldUtil {
    /**
     * 动态构造对象
     *
     * @param field
     * @param tClass
     * @param <T>
     * @return
     */
    public static <T extends BaseColumnProperty> T annotationToObject(Field field, Class<T> tClass) {
        T t = null;
        try {
            ExcelWriteProperty annotation = field.getAnnotation(ExcelWriteProperty.class);
            if (null != annotation && null != tClass) {
                String[] headValue = annotation.head();
                boolean ignoreCell = annotation.ignoreCell();
                int mergeCellIndex = annotation.mergeCellIndex();
                int mergeRowCount = annotation.mergeRowCount();
                String divideCellFormula = annotation.divideCellFormula();
                String[] sumCellFormula = annotation.sumCellFormula();
                String dateFormat = annotation.dateFormat();
                boolean isSeqNo = annotation.isSeqNo();
                Class<?>[] fieldClass = {Field.class, int.class, List.class, boolean.class, boolean.class, Integer.class, int.class, String.class, String.class, List.class};
                Object[] fieldValue = {field, annotation.index(),
                        Arrays.asList(headValue), isSeqNo, ignoreCell, mergeCellIndex >= 0 ? mergeCellIndex : null,
                        mergeRowCount < 0 ? 0 : mergeRowCount, dateFormat, divideCellFormula, Arrays.asList(sumCellFormula)};
                Constructor<T> constructor = tClass.getDeclaredConstructor(fieldClass);
                t = constructor.newInstance(fieldValue);
            }
        } catch (Exception e) {
            log.error("ExcelWriteProperty annotation cast to object fail! {}", e.getMessage());
            throw new ExcelParseException("ExcelWriteProperty annotation cast to object fail!");
        }
        return t;
    }

    /**
     * @param cls
     * @return
     */
    public static List<Field> getObjectField(Class<? extends BaseRowModel> cls) {
        List<Field> fieldList = new ArrayList<>();
        Class tempClass = cls;
        //判断是否为BaseRowModel 的第一个子类
        boolean isBaseRowModelFirstSub = BaseRowModel.class.equals(tempClass.getSuperclass());
        while (tempClass != null) {
            //直接继承 BaseRowModel 效率更快
            if (isBaseRowModelFirstSub) {
                fieldList.addAll(Arrays.asList(tempClass.getDeclaredFields()));
                break;
            } else {
                //只认子类的字段
                Field[] declaredFields = tempClass.getDeclaredFields();
                for (int i = 0; i < declaredFields.length; i++) {
                    Field declaredField = declaredFields[i];
                    boolean b = fieldList.stream().anyMatch(v -> v.getName().equals(declaredField.getName()));
                    if (!b) {
                        fieldList.add(declaredFields[i]);
                    }
                }
            }
            tempClass = tempClass.getSuperclass();
        }
        return fieldList;
    }

    public static <T> List<Field> getObjectAllField(Class<T> cls) {
        List<Field> fields = null;
        Map<String, Field> fieldAllMap = getObjectAllFieldMap(cls);
        if (fieldAllMap != null && fieldAllMap.size() > 0) {
            fields = fieldAllMap.values().stream().collect(Collectors.toList());
        }
        return fields;
    }

    public static <T> Map<String, Field> getObjectAllFieldMap(Class<T> cls) {
        Map<String, Field> fieldAllMap = new HashMap<>();
        Class<? super T> superclass = cls;
        do {
            if (superclass == null) {
                break;
            }
            Field[] declaredFields = superclass.getDeclaredFields();
            if (declaredFields != null && declaredFields.length > 0) {
                Map<String, Field> fieldMap = Arrays.stream(declaredFields).filter(v -> !"serialVersionUID".equals(v.getName()) && !fieldAllMap.containsKey(v.getName())).collect(Collectors.toMap(v -> v.getName(), Function.identity(), (k1, k2) -> k1));
                if (fieldMap != null) {
                    fieldAllMap.putAll(fieldMap);
                }
            }
            superclass = superclass.getSuperclass();
        } while (true);
        return fieldAllMap;
    }

    public static <T> void setValueToField(T instance, Field field, Object value) throws Exception {
        if (value == null || instance == null || field == null) {
            return;
        }
        field.setAccessible(true);
        Class<?> type = field.getType();
        if (type.equals(String.class)) {
            field.set(instance, value);
        } else if (type.equals(Double.class)) {
            field.set(instance, Double.parseDouble(value.toString()));
        } else if (type.equals(Integer.class)) {
            field.set(instance, Integer.valueOf(String.valueOf(value)));
        } else if (type.equals(Float.class)) {
            field.set(instance, Float.valueOf(value.toString()));
        } else if (type.equals(Boolean.class)) {
            field.set(instance, Boolean.valueOf(value.toString()));
        } else if (type.equals(Character.class)) {
            field.set(instance, (char) value);
        } else if (type.equals(Long.class)) {
            field.set(instance, Long.parseLong(value.toString()));
        } else if (type.equals(BigDecimal.class)) {
            field.set(instance, new BigDecimal(value.toString()));
        } else if (type.equals(Date.class)) {
            SimpleDateFormat sf = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss");
            Date date = sf.parse(value.toString());
            field.set(instance, date);
        }

    }
}
