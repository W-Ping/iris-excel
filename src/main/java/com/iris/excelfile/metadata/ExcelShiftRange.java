package com.iris.excelfile.metadata;

/**
 * 文档模板偏移范围
 *
 * @author liu_wp
 * @date Created in 2019/6/17 13:25
 * @see
 */
public class ExcelShiftRange {
    /**
     * 开始行
     */
    private int startRow;
    /**
     * 末尾行
     */
    private int endRow;
    /**
     * 移动行数
     */
    private int shiftNumber = 1;
    private boolean copyRowHeight;
    private boolean resetOriginalRowHeight;

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public int getShiftNumber() {
        return shiftNumber;
    }

    public void setShiftNumber(int shiftNumber) {
        if (shiftNumber <= 0) {
            shiftNumber = 1;
        }
        this.shiftNumber = shiftNumber;
    }

    public boolean isCopyRowHeight() {
        return copyRowHeight;
    }

    public void setCopyRowHeight(boolean copyRowHeight) {
        this.copyRowHeight = copyRowHeight;
    }

    public boolean isResetOriginalRowHeight() {
        return resetOriginalRowHeight;
    }

    public void setResetOriginalRowHeight(boolean resetOriginalRowHeight) {
        this.resetOriginalRowHeight = resetOriginalRowHeight;
    }
}
