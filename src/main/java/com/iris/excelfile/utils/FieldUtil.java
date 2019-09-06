package com.iris.excelfile.utils;

import com.iris.excelfile.annotation.ExcelReadProperty;
import com.iris.excelfile.annotation.ExcelReadSheet;
import com.iris.excelfile.annotation.ExcelWriteProperty;
import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.BaseColumnProperty;
import com.iris.excelfile.metadata.BaseRowModel;
import com.iris.excelfile.metadata.ExcelReadColumnProperty;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.formula.functions.T;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
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

    public static ExcelReadColumnProperty annotationReadToObject(Field field, Class<T> tClass) {
        ExcelReadColumnProperty t = new ExcelReadColumnProperty();
        try {
            ExcelReadProperty annotation = field.getAnnotation(ExcelReadProperty.class);
            if (null != annotation && null != tClass) {
                ExcelReadSheet excelReadSheet = tClass.getAnnotation(ExcelReadSheet.class);
                int sheetNo = excelReadSheet.SheetNo();
                int startRowIndex = excelReadSheet.startRowIndex();
                int startCellIndex = excelReadSheet.startCellIndex();
                int cellIndex = annotation.cellIndex();
                String dateFormat = annotation.dateFormat();
                String numberFormat = annotation.numberFormat();
                t.setCellIndex(cellIndex);
                t.setDateFormat(dateFormat);
                t.setNumberFormat(numberFormat);
                t.setSheetNo(sheetNo);
                t.setStartRowIndex(startRowIndex);
                t.setStartCellIndex(startCellIndex);
                t.setDataClass(tClass);
                t.setField(field);
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.error("ExcelReadColumnProperty annotation cast to object fail! {}", e.getMessage());
            throw new ExcelParseException("ExcelReadColumnProperty annotation cast to object fail!");
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

    public static <T> void setValueToField(T instance, Field field, Object value, String format) throws Exception {
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
            Date dateValue = (Date) value;
            Instant instant = dateValue.toInstant();
            ZoneId zoneId = ZoneId.systemDefault();
            LocalDateTime localDateTime = instant.atZone(zoneId).toLocalDateTime();
            format = StringUtils.isNotBlank(format) ? format : "yyyy-MM-dd HH:mm:ss";
            String date = localDateTime.format(DateTimeFormatter.ofPattern(format));
            dateValue = new Date(date);
//            dateValue = Date.from(zonedDateTime.toInstant());
//            if (dateValue.contains("-") || dateValue.contains("/") || dateValue.contains(":")) {
//                date = TypeUtil.getSimpleDateFormatDate(dateValue, format);
//            } else {
//                Double d = Double.parseDouble(dateValue);
//                dateValue = HSSFDateUtil.getJavaDate(d, true);
//            }
            int year = dateValue.getYear();
            int hours = dateValue.getHours();
            System.out.println(year + ":" + hours);
            if (dateValue != null) {
                field.set(instance, dateValue);
            } else {
                log.warn("[{}] date cost to Date is fail ", dateValue);
            }
        }

    }
}
