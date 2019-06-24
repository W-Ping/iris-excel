package com.iris.excelfile.metadata;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 17:01
 * @see
 */
public class ExcelColumnProperty extends BaseColumnProperty {

    private Field field;
    private boolean ignoreField;
    private int index = 99999;
    /**
     * 合并列索引
     */
    private Integer mergeCellIndex;
    /**
     * 合并行数
     */
    private int mergeRowCount;
    /**
     * 注解的表头信息
     */
    private List<String> head = new ArrayList<>();
    /**
     * excel sum公式
     */
    private List<String> sumCellFormula = new ArrayList<>();
    /**
     * excel 乘法公式
     */
    private String divideCellFormula;
    /**
     * excel 乘法公式
     */
    private String multiCellFormula;
    /**
     * 日期格式
     */
    private String dateFormat;
    /**
     * 是否为序列号
     */
    private boolean isSeqNo;

    public ExcelColumnProperty() {
    }

    public ExcelColumnProperty(Field field, int index, List<String> head) {
        this(field, index, head, null, null, null);
    }

    public ExcelColumnProperty(Field field, int index, List<String> head, String dateFormat, String divideCellFormula, List<String> sumCellFormula) {
        this(field, index, head, false, false, null, 0, dateFormat, divideCellFormula, sumCellFormula);
    }

    public ExcelColumnProperty(Field field, int index, List<String> head, boolean isSeqNo, boolean ignoreField, Integer mergeCellIndex, int mergeRowCount, String dateFormat, String divideCellFormula, List<String> sumCellFormula) {
        this.field = field;
        this.index = index;
        this.head = head;
        this.isSeqNo = isSeqNo;
        this.ignoreField = ignoreField;
        this.mergeCellIndex = mergeCellIndex;
        this.mergeRowCount = mergeRowCount;
        this.dateFormat = dateFormat;
        this.divideCellFormula = divideCellFormula;
        this.sumCellFormula = sumCellFormula;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public List<String> getHead() {
        return head;
    }

    public void setHead(List<String> head) {
        this.head = head;
    }

    public List<String> getSumCellFormula() {
        return sumCellFormula;
    }

    public void setSumCellFormula(List<String> sumCellFormula) {
        this.sumCellFormula = sumCellFormula;
    }

    public boolean isIgnoreField() {
        return ignoreField;
    }

    public void setIgnoreField(boolean ignoreField) {
        this.ignoreField = ignoreField;
    }

    public Integer getMergeCellIndex() {
        return mergeCellIndex;
    }

    public void setMergeCellIndex(Integer mergeCellIndex) {
        this.mergeCellIndex = mergeCellIndex;
    }

    public String getDivideCellFormula() {
        return divideCellFormula;
    }

    public void setDivideCellFormula(String divideCellFormula) {
        this.divideCellFormula = divideCellFormula;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public int getMergeRowCount() {
        return mergeRowCount;
    }

    public void setMergeRowCount(int mergeRowCount) {
        if (mergeRowCount < 0) {
            mergeRowCount = 0;
        }
        this.mergeRowCount = mergeRowCount;
    }

    public String getMultiCellFormula() {
        return multiCellFormula;
    }

    public void setMultiCellFormula(String multiCellFormula) {
        this.multiCellFormula = multiCellFormula;
    }

    @Override
    public int compareTo(ExcelColumnProperty o) {
        int x = this.index;
        int y = o.getIndex();
        return (x < y) ? -1 : ((x == y) ? 0 : 1);
    }

    public boolean isSeqNo() {
        return isSeqNo;
    }

    public void setSeqNo(boolean seqNo) {
        isSeqNo = seqNo;
    }
}
