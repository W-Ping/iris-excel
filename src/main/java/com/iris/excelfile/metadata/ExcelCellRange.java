package com.iris.excelfile.metadata;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 20:30
 * @see
 */
public class ExcelCellRange {

    /**
     * 开始行数
     */
    private int firstRowIndex = 1;
    private int lastRowIndex;
    /**
     * 开始列数
     */
    private int firstCellIndex = 1;
    private int lastCellIndex;

    public ExcelCellRange() {
    }

    public ExcelCellRange(int firstRowIndex, int lastRowIndex, int firstCellIndex, int lastCellIndex) {
        this.firstRowIndex = firstRowIndex;
        this.lastRowIndex = lastRowIndex;
        this.firstCellIndex = firstCellIndex;
        this.lastCellIndex = lastCellIndex;
    }

    public int getFirstRowIndex() {
        return firstRowIndex;
    }

    public void setFirstRowIndex(int firstRowIndex) {
        this.firstRowIndex = firstRowIndex;
    }

    public int getLastRowIndex() {
        return lastRowIndex;
    }

    public void setLastRowIndex(int lastRowIndex) {
        this.lastRowIndex = lastRowIndex;
    }

    public int getFirstCellIndex() {
        return firstCellIndex;
    }

    public void setFirstCellIndex(int firstCellIndex) {
        this.firstCellIndex = firstCellIndex;
    }

    public int getLastCellIndex() {
        return lastCellIndex;
    }

    public void setLastCellIndex(int lastCellIndex) {
        this.lastCellIndex = lastCellIndex;
    }
}
