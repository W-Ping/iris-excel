package com.iris.excelfile.metadata;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author liu_wp
 * @date Created in 2019/3/7 18:52
 * @see
 */
public class ExcelFont {
    /**
     * 字体名称 宋体
     */
    private String fontName;

    /**
     * 字体大小 14
     */
    private short fontHeightInPoints;

    /**
     * 是否加粗
     */
    private boolean bold;

    /**
     * 字体颜色
     */
    private IndexedColors fontColor;

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public short getFontHeightInPoints() {
        return fontHeightInPoints;
    }

    public void setFontHeightInPoints(short fontHeightInPoints) {
        this.fontHeightInPoints = fontHeightInPoints;
    }

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public IndexedColors getFontColor() {
        return fontColor;
    }

    public void setFontColor(IndexedColors fontColor) {
        this.fontColor = fontColor;
    }
}
