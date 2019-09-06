package com.iris.excelfile.core.builder.impl;

import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.context.ReadContext;
import com.iris.excelfile.core.builder.IReadBuilder;
import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.ExcelReadColumnProperty;
import com.iris.excelfile.utils.FieldUtil;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * @author liu_wp
 * @date 2019/8/30 9:21
 * @see
 */
public class ReadBuiderImpl implements IReadBuilder {
    private ReadContext readContext;
    private Map<Integer, List<ExcelReadColumnProperty>> excelSheetDataMap = new HashMap<>();
    private static final DataFormatter FORMATTER = new DataFormatter();

    public ReadBuiderImpl(InputStream inputStream, ExcelTypeEnum excelType) {
        this.readContext = new ReadContext(inputStream, excelType);
    }

    @Override
    public void initDataClass(List<Class> classes) {
        excelSheetDataMap.clear();
        for (Class aClass : classes) {
            List<Field> objectAllField = FieldUtil.getObjectAllField(aClass);
            if (!CollectionUtils.isEmpty(objectAllField)) {
                List<ExcelReadColumnProperty> excelSheetData = new ArrayList<>();
                ExcelReadColumnProperty t = null;
                for (Field field : objectAllField) {
                    t = FieldUtil.annotationReadToObject(field, aClass);
                    excelSheetData.add(t);
                }
                if (!CollectionUtils.isEmpty(excelSheetData)) {
                    excelSheetDataMap.put(excelSheetData.get(0).getSheetNo(), excelSheetData);
                }
            }

        }
    }

    @Override
    public Map<Integer, Object> addExcelToObject() {
        Workbook workbook = readContext.getWorkbook();
        int activeSheetIndex = workbook.getActiveSheetIndex();
        if (activeSheetIndex != excelSheetDataMap.size()) {
            throw new ExcelParseException("excel sheet map object is unequally");
        }
        Map<Integer, Object> map = new HashMap<>();
        for (int i = 0; i < activeSheetIndex; i++) {
            Sheet sheetAt = workbook.getSheetAt(i);
            List<ExcelReadColumnProperty> excelSheetDataList = excelSheetDataMap.get(i);
            if (!CollectionUtils.isEmpty(excelSheetDataList)) {
                Object data = null;
                for (int j = 0; j < excelSheetDataList.size(); j++) {
                    data = addExcelToObject(sheetAt, data, excelSheetDataList.get(j), i);
                }
                if (data != null) {
                    map.put(i, data);
                }
            }
        }
        return map;
    }

    @Override
    public Object addOneSheetExcelToObject() {
        Workbook workbook = readContext.getWorkbook();
        Integer sheetNo = excelSheetDataMap.keySet().iterator().next();
        if (sheetNo == null) {
            throw new ExcelParseException("excel sheet is not exist");
        }
        List<ExcelReadColumnProperty> excelSheetDataList = excelSheetDataMap.get(sheetNo);
        Sheet sheetAt = workbook.getSheetAt(sheetNo);
        Object data = null;
        if (!CollectionUtils.isEmpty(excelSheetDataList)) {
            for (int j = 0; j < excelSheetDataList.size(); j++) {
                data = addExcelToObject(sheetAt, data, excelSheetDataList.get(j), sheetNo);
            }
        }
        return data;
    }

    private Object addExcelToObject(Sheet sheet, Object instance, ExcelReadColumnProperty excelReadColumnProperty, int sheetNo) {
        if (sheetNo == excelReadColumnProperty.getSheetNo()) {
            try {
                if (instance == null) {
                    instance = excelReadColumnProperty.getDataClass().newInstance();
                }
                int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
                int startRowIndex = excelReadColumnProperty.getStartRowIndex();
                if (startRowIndex <= physicalNumberOfRows) {
                    for (int rowIndex = startRowIndex; rowIndex < physicalNumberOfRows; rowIndex++) {
                        Row row = sheet.getRow(rowIndex);
                        int physicalNumberOfCells = row.getPhysicalNumberOfCells();
                        int startCellIndex = excelReadColumnProperty.getStartCellIndex();
                        if (startCellIndex <= physicalNumberOfCells)
                            for (int cellIndex = startCellIndex; cellIndex < physicalNumberOfCells; cellIndex++) {
                                if (cellIndex == excelReadColumnProperty.getCellIndex()) {
                                    addCellDataToField(row.getCell(cellIndex), excelReadColumnProperty.getField(), instance, excelReadColumnProperty.getDateFormat());
                                }
                            }
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return instance;
    }

    private void addCellDataToField(Cell cell, Field field, Object instance, String dateFormat) {
        try {
            Object value = formatCellValue(cell);
            if (instance == null || field == null || value == null) {
                return;
            }
//            System.out.println(field.getName() + ":" + field.getType().getSimpleName() + ":" + field.getType());
//            field.set(instance, value);
            FieldUtil.setValueToField(instance, field, value, dateFormat);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Object formatCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        int cellType = cell.getCellType();
        Object value = null;
        if (cellType == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(cell)) {
//            value = getFormattedDateString(cell);
            value = HSSFDateUtil.getJavaDate(cell.getNumericCellValue(), false);
        } else {
            value = FORMATTER.formatCellValue(cell);
        }
        return value;
    }

    private String performDateFormatting(Date d, Format dateFormat) {
        if (dateFormat != null) {
            return dateFormat.format(d);
        }
        return d.toString();
    }

    private String getFormattedDateString(Cell cell) {
        Format dateFormat = FORMATTER.createFormat(cell);
        if (dateFormat instanceof ExcelStyleDateFormatter) {
            // Hint about the raw excel value
            ((ExcelStyleDateFormatter) dateFormat).setDateToBeFormatted(
                    cell.getNumericCellValue()
            );
        }
        Date d = cell.getDateCellValue();
        return performDateFormatting(d, dateFormat);
    }
}
