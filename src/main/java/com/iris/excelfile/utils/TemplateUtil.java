package com.iris.excelfile.utils;

import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.ExcelFont;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * 模板
 */
@Slf4j
public class TemplateUtil {
    /**
     * @param workbook
     * @param sheet
     * @param excelFont
     * @param rowIndex
     * @param cellIndex
     * @param startIndex  替换开始索引 负数从右边开始替换
     * @param size        替换的长度
     * @param refMap
     * @param replaceKeys
     */
    public static void replaceTemplate(Workbook workbook, Sheet sheet, ExcelFont excelFont, int rowIndex, int cellIndex, int startIndex, int size, Map<String, String> refMap, String... replaceKeys) {
        Cell cell = getCell(getRow(sheet, rowIndex), cellIndex);
        String stringCellValue = replaceTemplate(cell.getStringCellValue(), refMap, replaceKeys);
        Integer[] indexArr = calculateIndex(stringCellValue, startIndex, size);
        StyleUtil.buildTextStyle(stringCellValue, workbook, cell, excelFont, indexArr[0], indexArr[1]);
    }

    private static Integer[] calculateIndex(String stringCellValue, int startIndex, int size) {
        int endIndex = 0;
        if (startIndex < 0) {
            endIndex = stringCellValue.length() + startIndex + 1;
            endIndex = endIndex < 0 ? 0 : endIndex;
            startIndex = endIndex - Math.abs(size);
            startIndex = startIndex < 0 ? 0 : startIndex;
        } else {
            endIndex = startIndex + Math.abs(size);
            endIndex = endIndex > stringCellValue.length() - 1 ? stringCellValue.length() - 1 : endIndex;
        }
        Integer[] indexArr = new Integer[2];
        indexArr[0] = startIndex;
        indexArr[1] = endIndex;
        return indexArr;
    }

    /**
     * @param workbook
     * @param sheet
     * @param excelFont
     * @param rowIndex
     * @param cellIndex
     * @param refMap
     * @param replaceKeys
     */
    public static void replaceTemplate(Workbook workbook, Sheet sheet, ExcelFont excelFont, int rowIndex, int cellIndex, Map<String, String> refMap, String... replaceKeys) {
        Cell cell = getCell(getRow(sheet, rowIndex), cellIndex);
        String stringCellValue = replaceTemplate(cell.getStringCellValue(), refMap, replaceKeys);
        replaceTemplate(workbook, sheet, excelFont, rowIndex, cellIndex, 0, stringCellValue.length(), refMap, replaceKeys);
    }

    /**
     * @param workbook
     * @param cell
     * @param excelFont
     * @param refMap
     * @param replaceKeys
     */
    public static void replaceTemplate(Workbook workbook, Cell cell, ExcelFont excelFont, Map<String, String> refMap, String... replaceKeys) {
        String stringCellValue = replaceTemplate(cell.getStringCellValue(), refMap, replaceKeys);
        StyleUtil.buildTextStyle(stringCellValue, workbook, cell, excelFont, 0, stringCellValue.length());
    }

    /**
     * @param stringCellValue
     * @param refMap
     * @param replaceKeys
     * @return
     */
    public static String replaceTemplate(String stringCellValue, Map<String, String> refMap, String... replaceKeys) {
        for (int i = 0; i < replaceKeys.length; i++) {
            if (StringUtils.isNotBlank(refMap.get(replaceKeys[i]))) {
                //模板避免有特殊字符
                stringCellValue = stringCellValue.replace(replaceKeys[i], refMap.get(replaceKeys[i]));
            }
        }
        return stringCellValue;
    }

    public static Row getRow(Sheet sheet, int rowIndex) {
        Row row = null;
        try {
            row = sheet.getRow(rowIndex);
            if (row == null) {
                throw new ExcelParseException("【" + rowIndex + "】行不存在");
            }
        } catch (Exception e) {
            log.error("获取模板数据错误{}", e.getMessage());
            throw new ExcelParseException("获取模板数据错误");
        }
        return row;
    }

    public static Cell getCell(Row row, int cellIndex) {
        Cell cell = null;
        try {
            cell = row.getCell(cellIndex);
            if (cell == null) {
                throw new ExcelParseException("【" + cellIndex + "】列不存在");
            }
        } catch (Exception e) {
            log.error("获取模板数据错误{}", e.getMessage());
            throw new ExcelParseException("获取模板数据错误");
        }
        return cell;
    }
}
