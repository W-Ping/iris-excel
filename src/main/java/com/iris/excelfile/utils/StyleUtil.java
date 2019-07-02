package com.iris.excelfile.utils;


import com.iris.excelfile.constant.BorderEnum;
import com.iris.excelfile.constant.DataFormatEnum;
import com.iris.excelfile.constant.FileConstant;
import com.iris.excelfile.constant.TableBodyEnum;
import com.iris.excelfile.metadata.ExcelFont;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.HashMap;
import java.util.Map;

/**
 * @author liu_wp
 * @date Created in 2019/3/8 10:29
 * @see
 */
public class StyleUtil {

    private static CellStyle fieldNullCellStyle = null;
    public static Map<Integer, CellStyle> defaultContentCellStyle = new HashMap<>();
    public static Map<Integer, CellStyle> defaultHeadStyle = new HashMap<>();
    public static Map<Integer, CellStyle> defaultFootCellStyle = new HashMap<>();


    /**
     * 清空样式缓存 避免错误
     * this Style does not belong to the supplied Workbook Stlyes Source
     */
    public static void clearDefaultStyleCache() {
        fieldNullCellStyle = null;
        defaultContentCellStyle.clear();
        defaultHeadStyle.clear();
        defaultFootCellStyle.clear();
    }

    /**
     * 表头默认样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle defaultHeadStyle(Workbook workbook, int tableNo, boolean locked) {
        CellStyle cellStyle = defaultHeadStyle.get(tableNo);
        if (cellStyle == null) {
            cellStyle = buildBaseIsBoldCellStyle(workbook, FileConstant.DEFAULT_HEAD_FONT_SIZE, IndexedColors.GREY_50_PERCENT, locked);
            defaultHeadStyle.put(tableNo, cellStyle);
        }
        return cellStyle;
    }

    /**
     * 内容默认样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle defaultContentCellStyle(Workbook workbook, int tableNo, IndexedColors defaultColor, boolean locked) {
        CellStyle cellStyle = defaultContentCellStyle.get(tableNo);
        if (cellStyle == null) {
            cellStyle = buildBaseIsNotBoldCellStyle(workbook, FileConstant.DEFAULT_CONTENT_FONT_SIZE, defaultColor != null ? defaultColor : IndexedColors.WHITE, locked);
            defaultContentCellStyle.put(tableNo, cellStyle);
        }
        return cellStyle;
    }

    /**
     * 表尾默认样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle defaultFootCellStyle(Workbook workbook, int tableNo, boolean locked) {
        CellStyle cellStyle = defaultFootCellStyle.get(tableNo);
        if (cellStyle == null) {
            cellStyle = buildBaseIsBoldCellStyle(workbook, FileConstant.DEFAULT_FOOT_FONT_SIZE, IndexedColors.GREY_25_PERCENT, locked);
            defaultFootCellStyle.put(tableNo, cellStyle);
        }
        return cellStyle;
    }

    /**
     * @param workbook
     * @param isFieldNull
     * @return
     */
    public static CellStyle buildFieldIsNullStyle(Workbook workbook, boolean isFieldNull) {
        if (isFieldNull) {
            fieldNullCellStyle = fieldNullCellStyle != null ? fieldNullCellStyle : buildBaseIsNotBoldCellStyle(workbook, FileConstant.DEFAULT_CONTENT_FONT_SIZE, IndexedColors.WHITE, false);
            fieldNullCellStyle.setBorderBottom((short) 0);
            fieldNullCellStyle.setBorderTop((short) 0);
            return fieldNullCellStyle;
        }
        return null;
    }

    /**
     * @param currentSheet
     * @param sheetWidthMap
     * @return
     */
    public static Sheet buildTableWidthStyle(Sheet currentSheet, Map<Integer, Integer> sheetWidthMap) {
        currentSheet.setDefaultColumnWidth(FileConstant.DEFAULT_COLUMN_WIDTH);
        for (Map.Entry<Integer, Integer> entry : sheetWidthMap.entrySet()) {
            if (entry.getKey() == null || entry.getValue() == null) {
                continue;
            }
            currentSheet.setColumnWidth(entry.getKey(), (entry.getValue().intValue() * 256));
        }
        return currentSheet;
    }

    /**
     * @param cellStyle
     * @param index
     * @param bol
     * @param borderEnum
     */
    public static void buildCellBorderStyle(CellStyle cellStyle, int index, boolean bol, BorderEnum borderEnum) {
        if (cellStyle != null && bol) {
            short borderTop = cellStyle.getBorderBottom();
            //V.3.17
            // BorderStyle borderBottomEnum = cellStyle.getBorderBottomEnum();
            if (borderTop > 0) {
                if (cellStyle.getBorderTop() > 0) {
                    borderTop = cellStyle.getBorderTop();
                } else if (cellStyle.getBorderLeft() > 0) {
                    borderTop = cellStyle.getBorderLeft();
                } else {
                    borderTop = cellStyle.getBorderRight();
                }
            }
            switch (borderEnum) {
                case TOP:
                    cellStyle.setBorderTop(borderTop);
                    break;
                case LEFT:
                    cellStyle.setBorderLeft(borderTop);
                    break;
                case BOTTOM:
                    cellStyle.setBorderBottom(borderTop);
                    break;
                case RIGHT:
                    cellStyle.setBorderRight(borderTop);
                    break;
                default:
                    break;
            }
        }
    }


    /**
     * 设置背景颜色，字体，保护锁
     *
     * @param workbook
     * @param f
     * @param indexedColors
     * @return
     */
    public static CellStyle buildCellStyle(Workbook workbook, int tableNo, ExcelFont f,
                                           IndexedColors indexedColors, TableBodyEnum tableBodyEnum, boolean locked) {
        CellStyle cellStyle = null;
        if (TableBodyEnum.HEAD.equals(tableBodyEnum)) {
            cellStyle = defaultHeadStyle(workbook, tableNo, locked);
        } else if (TableBodyEnum.FOOT.equals(tableBodyEnum)) {
            cellStyle = defaultFootCellStyle(workbook, tableNo, locked);
        } else {
            cellStyle = defaultContentCellStyle(workbook, tableNo, indexedColors, locked);
        }
        if (f != null) {
            Font font = workbook.createFont();
            font.setFontName(f.getFontName());
            font.setFontHeightInPoints(f.getFontHeightInPoints());
            font.setBold(f.isBold());
            if (f.getFontColor() != null) {
                font.setColor(f.getFontColor().getIndex());
            }
            cellStyle.setFont(font);
        }
        if (indexedColors != null) {
            cellStyle.setFillForegroundColor(indexedColors.getIndex());
        }
        return cellStyle;
    }

    /**
     * @param workbook
     * @param cellStyle
     * @param dataFormatEnum
     * @param isFormat
     */
    public static void buildAllCellDataFormatStyle(Workbook workbook, CellStyle cellStyle, DataFormatEnum dataFormatEnum, boolean isFormat) {
        if (!isFormat || dataFormatEnum == null || DataFormatEnum.NORMAL.equals(dataFormatEnum)) {
            return;
        }
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(dataFormatEnum.getFormat()));
    }

    public static CellStyle buildBaseIsBoldCellStyle(Workbook workbook, Short fontSize, IndexedColors indexedColor, boolean locked) {
        return buildBaseCellStyle(workbook, fontSize, indexedColor, true, locked);
    }

    public static CellStyle buildBaseIsNotBoldCellStyle(Workbook workbook, Short fontSize, IndexedColors indexedColor, boolean locked) {
        return buildBaseCellStyle(workbook, fontSize, indexedColor, false, locked);
    }

    public static CellStyle buildBaseCellStyle(Workbook workbook, Short fontSize, IndexedColors indexedColor, boolean bold, boolean locked) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName(FileConstant.DEFAULT_FONT_NAME);
        font.setFontHeightInPoints(fontSize == null || fontSize <= 0 ? FileConstant.DEFAULT_CONTENT_FONT_SIZE : fontSize);
        font.setBold(bold);
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        cellStyle.setLocked(locked);
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setFillForegroundColor(indexedColor != null ? indexedColor.getIndex() : IndexedColors.WHITE.getIndex());
        return cellStyle;
    }

    public static void buildTextStyle(String text, Workbook workbook, Cell cell, ExcelFont font) {
        buildTextStyle(text, workbook, cell, font, 0, text.length());
    }

    public static void buildTextStyle(String text, Workbook workbook, Cell cell, ExcelFont excelFont, int startIndex, int endIndex) {
        if (excelFont != null) {
            XSSFRichTextString richString = new XSSFRichTextString(text);
            Font font = workbook.createFont();
            font.setBold(excelFont.isBold());
            if (StringUtils.isNotBlank(font.getFontName())) {
                font.setFontName(font.getFontName());
            }
            if (excelFont.getFontColor() != null) {
                font.setColor(excelFont.getFontColor().getIndex());
            }
            if (excelFont.getFontHeightInPoints() > 0) {
                font.setFontHeightInPoints(excelFont.getFontHeightInPoints());
            }
            richString.applyFont(startIndex, endIndex, font);
            cell.setCellValue(richString);
        }
    }

}
