package com.iris.excelfile.metadata;


import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author liu_wp
 * @date Created in 2019/3/5 18:53
 * @see
 */
public class ExcelStyle {
    /**
     * 字体
     */
    private ExcelFont excelFont;
    private IndexedColors excelBackGroundColor;
    /**
     * 样式
     */
    private CellStyle currentCellStyle;
    /**
     * 默认样式
     */
    private CellStyle defaultCellStyle;

    public ExcelFont getExcelFont() {
        return excelFont;
    }

    public void setExcelFont(ExcelFont excelFont) {
        this.excelFont = excelFont;
    }

    public IndexedColors getExcelBackGroundColor() {
        return excelBackGroundColor;
    }

    public void setExcelBackGroundColor(IndexedColors excelBackGroundColor) {
        this.excelBackGroundColor = excelBackGroundColor;
    }

    public CellStyle getCurrentCellStyle() {
        return currentCellStyle;
    }

    public void setCurrentCellStyle(CellStyle currentCellStyle) {
        this.currentCellStyle = currentCellStyle;
    }

    public CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    public void setDefaultCellStyle(CellStyle defaultCellStyle) {
        this.defaultCellStyle = defaultCellStyle;
    }
}
