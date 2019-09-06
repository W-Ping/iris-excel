package com.iris.excelfile.metadata;

import org.apache.poi.ss.formula.functions.T;

import java.lang.reflect.Field;

/**
 * @author liu_wp
 * @date 2019/8/30 10:19
 * @see
 */
public class ExcelReadColumnProperty extends BaseColumnProperty {

    private Field field;

    private Class dataClass;
    /**
     *
     */
    private int sheetNo;
    /**
     *
     */
    private int cellIndex;

    /**
     *
     */
    private int startRowIndex;
    /**
     *
     */
    private int startCellIndex;
    /**
     * 日期格式
     */
    private String dateFormat;
    /**
     * 数字格式
     */
    private String numberFormat;

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public int getSheetNo() {
        return sheetNo;
    }

    public void setSheetNo(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public int getStartRowIndex() {
        return startRowIndex;
    }

    public void setStartRowIndex(int startRowIndex) {
        this.startRowIndex = startRowIndex;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public String getNumberFormat() {
        return numberFormat;
    }

    public void setNumberFormat(String numberFormat) {
        this.numberFormat = numberFormat;
    }

    public int getStartCellIndex() {
        return startCellIndex;
    }

    public void setStartCellIndex(int startCellIndex) {
        this.startCellIndex = startCellIndex;
    }

    public Class getDataClass() {
        return dataClass;
    }

    public void setDataClass(Class dataClass) {
        this.dataClass = dataClass;
    }
}
