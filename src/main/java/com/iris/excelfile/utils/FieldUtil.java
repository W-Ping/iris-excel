package com.iris.excelfile.utils;

import com.iris.excelfile.annotation.ExcelWriteProperty;
import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.BaseColumnProperty;
import com.iris.excelfile.metadata.BaseRowModel;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

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
                boolean ignore = annotation.ignore();
                int mergeCellIndex = annotation.mergeCellIndex();
                int mergeRowCount = annotation.mergeRowCount();
                String divideCellFormula = annotation.divideCellFormula();
                String[] sumCellFormula = annotation.sumCellFormula();
                String dateFormat = annotation.dateFormat();
                boolean isSeqNo = annotation.isSeqNo();
                Class<?>[] fieldClass = {Field.class, int.class, List.class, boolean.class, boolean.class, Integer.class, int.class, String.class, String.class, List.class};
                Object[] fieldValue = {field, annotation.index(),
                        Arrays.asList(headValue), isSeqNo, ignore, mergeCellIndex >= 0 ? mergeCellIndex : null,
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
}
