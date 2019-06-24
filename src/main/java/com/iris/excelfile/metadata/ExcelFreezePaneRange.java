package com.iris.excelfile.metadata;

/**
 * 冻结范围参数
 */
public class ExcelFreezePaneRange {
    /**
     * 冻结的列数 (从excel展示列开始)
     */
    private int cellCount;
    /**
     * 冻结的行数
     */
    private int rowCount;
    /**
     * 右边区域[可见]的首列序号
     */
    private int cellIndex;
    /**
     * 下边区域[可见]的首行序号
     */
    private int rowIndex;

    public int getCellCount() {
        return cellCount;
    }

    public void setCellCount(int cellCount) {
        this.cellCount = cellCount;
    }

    public int getRowCount() {
        return rowCount;
    }

    public void setRowCount(int rowCount) {
        this.rowCount = rowCount;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }
}
