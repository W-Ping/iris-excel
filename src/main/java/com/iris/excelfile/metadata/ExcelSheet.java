package com.iris.excelfile.metadata;


import com.iris.excelfile.constant.DataFormatEnum;
import com.iris.excelfile.constant.TableLayoutEnum;
import com.iris.excelfile.core.handler.extend.IWriteLoadTemplateHandler;
import com.iris.excelfile.exception.ExcelParseException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.*;

/**
 * @author liu_wp
 * @date Created in 2019/3/5 12:12
 * @see
 */
@Slf4j
public class ExcelSheet {
    private int sheetNo;
    private String sheetName;
    /**
     * 默认自动布局
     */
    private boolean isAutoLayOut = true;
    /**
     * table
     */
    private List<ExcelTable> excelTables;
    /**
     * 列的宽度
     */
    private Map<Integer, Integer> columnWidthMap = new HashMap<>();
    /**
     * 行的高度
     */
    private Integer heightInPoints;
    /**
     * 合并格式
     * Integer[]={startRow,lastRow,startCell,endCell}
     */
    private List<Integer[]> mergeData;
    /**
     * 默认数据格式
     */
    private DataFormatEnum sheetDataDefaultFormat;
    /**
     * 模板加载之前handler
     */
    private IWriteLoadTemplateHandler writeLoadTemplateHandler;
    /**
     * sheet 冻结范围
     */
    private ExcelFreezePaneRange freezePaneRange;
    /**
     * 文档保护密码 默认123
     */
    private String protectSheetPwd;
    /**
     * 文档解除保护需要文档保护密码
     */
    private boolean locked;
    /**
     * table 缓存
     */
    private Map<Integer, ExcelTable> sheetTablesCache = new LinkedHashMap<>();


    /**
     * 默认背景颜色
     */
    private IndexedColors defaultBackGroundColor;

    public ExcelSheet(int sheetNo, List<ExcelTable> excelTables) {
        this(sheetNo, null, excelTables, true);
    }

    public ExcelSheet(int sheetNo, String sheetName, List<ExcelTable> excelTables) {
        this(sheetNo, sheetName, excelTables, true);
    }

    public ExcelSheet(int sheetNo, List<ExcelTable> excelTables, boolean isAutoLayOut) {
        this(sheetNo, null, excelTables, isAutoLayOut);
    }

    public ExcelSheet(int sheetNo, String sheetName, List<ExcelTable> excelTables, boolean isAutoLayOut) {
        this.sheetName = sheetName;
        this.sheetNo = sheetNo >= 0 ? sheetNo : 0;
        this.excelTables = excelTables;
        this.isAutoLayOut = isAutoLayOut;
    }


    /**
     * 初始化计算table布局参数和table样式
     */
    public void initCurrentSheetTables(ExcelSheet excelSheet) {
        if (!CollectionUtils.isEmpty(excelTables) && excelTables.size() >= 1) {
            Collections.sort(this.excelTables);
            log.info(" sheet tables is auto layout");
            for (int i = 0; i < this.excelTables.size(); i++) {
                ExcelTable excelTable = excelTables.get(i);
                initTableProperties(excelSheet, excelTable);
                //计算布局
                if (isAutoLayOut && excelTables.size() > 1) {
                    calculateSheetTableLayOut(excelTable, i);
                }
                sheetTablesCache.put(excelTable.getTableNo(), excelTable);
            }
        }
    }

    /**
     * 计算sheet里table 布局范围
     *
     * @param currTable
     * @param i
     */
    public void calculateSheetTableLayOut(ExcelTable currTable, int i) {
        ExcelCellRange currTableCellRange = currTable.getTableCellRange();
        if (i > 0) {
            TableLayoutEnum tableLayoutEnum = currTable.getTableLayoutEnum();
            ExcelTable maxRowTable = getMaxRowTable(i - 1);
            int rowNum = currTableCellRange.getLastRowIndex() - currTableCellRange.getFirstRowIndex();
            int headRowNum = currTable.isNeedHead() ? currTable.getExcelHeadProperty().getHeadRowNum() : 0;
            if (TableLayoutEnum.RIGHT.equals(tableLayoutEnum)) {
                ExcelTable preTable = this.excelTables.get(i - 1);
                ExcelCellRange preTableCellRange = preTable.getTableCellRange();
                currTableCellRange.setFirstCellIndex(preTableCellRange.getLastCellIndex() + currTable.getSpaceNum());
                currTableCellRange.setLastCellIndex(currTableCellRange.getLastCellIndex() + currTableCellRange.getFirstCellIndex());
                currTableCellRange.setFirstRowIndex(preTableCellRange.getFirstRowIndex());
                currTableCellRange.setLastRowIndex(currTableCellRange.getFirstRowIndex() + rowNum);
                currTable.setFirstCellIndex(currTableCellRange.getFirstCellIndex());
                currTable.setFirstRowIndex(currTableCellRange.getFirstRowIndex());
            } else if (TableLayoutEnum.BOTTOM.equals(tableLayoutEnum)) {
                currTableCellRange.setFirstRowIndex(maxRowTable.getTableCellRange().getLastRowIndex() + 1 + currTable.getSpaceNum());
                currTableCellRange.setLastRowIndex(currTableCellRange.getFirstRowIndex() + rowNum);
                currTable.setFirstRowIndex(currTableCellRange.getFirstRowIndex());
            } else {
                log.warn("文档布局只支持right和bottom");
                return;
            }
            currTable.setStartContentRowIndex(currTableCellRange.getFirstRowIndex() + headRowNum);
            currTable.setTableCellRange(currTableCellRange);
        }
        log.info("tableNo{} set TableCellRange 【firstRowIndex:{},lastRowIndex:{},firstCellIndex:{},lastCellIndex:{}】", currTable.getTableNo(),
                currTableCellRange.getFirstRowIndex(), currTableCellRange.getLastRowIndex(), currTableCellRange.getFirstCellIndex(), currTableCellRange.getLastCellIndex());
    }

    private void initTableProperties(ExcelSheet excelSheet, ExcelTable table) {
        initExcelHeadProperty(table);
        initTableRange(table);
        setSheetParams(excelSheet, table);
    }

    /**
     * @param excelSheet
     * @param excelTable
     */
    private void setSheetParams(ExcelSheet excelSheet, ExcelTable excelTable) {
        //默认数据格式
        if (excelTable.getTableDataDefaultFormat() == null) {
            excelTable.setTableDataDefaultFormat(excelSheet.getSheetDataDefaultFormat());
        }
        //默认背景颜色
        ExcelStyle tableStyle = excelTable.getTableStyle();
        if (tableStyle == null) {
            tableStyle = new ExcelStyle();
            tableStyle.setExcelBackGroundColor(excelSheet.getDefaultBackGroundColor());
        } else {
            if (tableStyle.getExcelBackGroundColor() == null) {
                tableStyle.setExcelBackGroundColor(excelSheet.getDefaultBackGroundColor());
            }
        }
        excelTable.setTableStyle(tableStyle);
        //行高
        excelTable.setHeightInPoints(excelSheet.getHeightInPoints());
        //是否上保护锁
        excelTable.setLocked(excelTable.isLocked() ? true : excelSheet.isLocked());
    }

    private void initExcelHeadProperty(ExcelTable table) {
        table.setExcelHeadProperty(new ExcelHeadProperty(table.getHeadClass(), table.getHead()));
        if (CollectionUtils.isEmpty(table.getHead())) {
            table.setHead(table.getExcelHeadProperty().getHead());
        }
    }

    private void initTableRange(ExcelTable table) {
        int headRowNum = table.isNeedHead() ? table.getExcelHeadProperty().getHeadRowNum() : 0;
        if (table.getStartContentRowIndex() <= 0) {
            table.setStartContentRowIndex(table.getFirstRowIndex() + headRowNum);
        }
        int tableDataIndex = 0;
        int minDataRowCount = table.getMinDataRowCount();
        if (!CollectionUtils.isEmpty(table.getData()) && table.getData().size() > minDataRowCount) {
            tableDataIndex = table.getData().size() - 1;
        } else if (minDataRowCount >= 1) {
            tableDataIndex = table.getMinDataRowCount() - 1;
        }
//        int lastRow = table.getStartContentRowIndex() + (!CollectionUtils.isEmpty(table.getData()) ? table.getData().size() - 1 : table.getMinDataRowCount() >= 1 ? table.getMinDataRowCount() - 1 : 0);
        int lastRow = table.getStartContentRowIndex() + tableDataIndex;
        table.setTableCellRange(new ExcelCellRange(table.getFirstRowIndex(), lastRow, table.getFirstCellIndex(), table.getHead().size() - 1 + table.getFirstCellIndex()));

    }

    private ExcelTable getMaxCellTable(int num) {
        ExcelTable excelTable = null;
        if (num >= this.excelTables.size()) {
            num = this.excelTables.size() - 1;
        }
        for (int j = num; j >= 0; j--) {
            ExcelTable currTable = this.excelTables.get(j);
            if (!TableLayoutEnum.RIGHT.equals(currTable.getTableLayoutEnum())) {
                if (excelTable == null) {
                    excelTable = currTable;
                }
                break;
            }
            if (j >= 1) {
                ExcelTable preTable = this.excelTables.get(j - 1);
                if (excelTable != null) {
                    currTable = excelTable;
                }
                if (currIsMax(currTable.getTableCellRange().getLastCellIndex(), preTable.getTableCellRange().getLastCellIndex())) {
                    excelTable = currTable;
                } else {
                    excelTable = preTable;
                }
            }
        }
        return excelTable;
    }

    private ExcelTable getMaxRowTable(int num) {
        ExcelTable excelTable = null;
        if (num >= this.excelTables.size()) {
            num = this.excelTables.size() - 1;
        }
        for (int j = num; j >= 0; j--) {
            ExcelTable currTable = this.excelTables.get(j);
            if (j >= 1) {
                ExcelTable preTable = this.excelTables.get(j - 1);
                if (excelTable != null) {
                    currTable = excelTable;
                }
                if (currIsMax(currTable.getTableCellRange().getLastRowIndex(), preTable.getTableCellRange().getLastRowIndex())) {
                    excelTable = currTable;
                } else {
                    excelTable = preTable;
                }
            }
        }
        if (excelTable == null) {
            excelTable = this.excelTables.get(0);
        }
        return excelTable;
    }

    private boolean currIsMax(int a, int b) {
        if (a > b) {
            return true;
        }
        return false;
    }


    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getSheetNo() {
        return sheetNo;
    }

    public void setSheetNo(int sheetNo) {
        this.sheetNo = sheetNo >= 0 ? sheetNo : 0;
    }

    public List<ExcelTable> getExcelTables() {
        return excelTables;
    }

    public void setExcelTables(List<ExcelTable> excelTables) {
        this.excelTables = excelTables;
    }

    public Map<Integer, Integer> getColumnWidthMap() {
        return columnWidthMap;
    }

    public void setColumnWidthMap(Map<Integer, Integer> columnWidthMap) {
        this.columnWidthMap = columnWidthMap;
    }

    public boolean isAutoLayOut() {
        return isAutoLayOut;
    }

    public void setAutoLayOut(boolean autoLayOut) {
        isAutoLayOut = autoLayOut;
    }

    public List<Integer[]> getMergeData() {
        return mergeData;
    }

    public DataFormatEnum getSheetDataDefaultFormat() {
        return sheetDataDefaultFormat;
    }

    public void setSheetDataDefaultFormat(DataFormatEnum sheetDataDefaultFormat) {
        this.sheetDataDefaultFormat = sheetDataDefaultFormat;
    }

    public IWriteLoadTemplateHandler getWriteLoadTemplateHandler() {
        return writeLoadTemplateHandler;
    }

    public void setWriteLoadTemplateHandler(IWriteLoadTemplateHandler writeLoadTemplateHandler) {
        this.writeLoadTemplateHandler = writeLoadTemplateHandler;
    }

    public ExcelFreezePaneRange getFreezePaneRange() {
        return freezePaneRange;
    }

    public void setFreezePaneRange(ExcelFreezePaneRange freezePaneRange) {
        this.freezePaneRange = freezePaneRange;
    }

    public Map<Integer, ExcelTable> getSheetTablesCache() {
        return sheetTablesCache;
    }

    public void setSheetTablesCache(Map<Integer, ExcelTable> sheetTablesCache) {
        this.sheetTablesCache = sheetTablesCache;
    }

    public Integer getHeightInPoints() {
        return heightInPoints;
    }

    public String getProtectSheetPwd() {
        return protectSheetPwd;
    }

    public void setProtectSheetPwd(String protectSheetPwd) {
        this.protectSheetPwd = protectSheetPwd;
    }

    public boolean isLocked() {
        return locked;
    }

    public void setLocked(boolean locked) {
        this.locked = locked;
    }

    public void setHeightInPoints(Integer heightInPoints) {
        this.heightInPoints = heightInPoints;
    }

    public IndexedColors getDefaultBackGroundColor() {
        return defaultBackGroundColor;
    }

    public void setDefaultBackGroundColor(IndexedColors defaultBackGroundColor) {
        this.defaultBackGroundColor = defaultBackGroundColor;
    }

    public void setMergeData(List<Integer[]> mergeData) {
        if (mergeData != null && mergeData.stream().filter(v -> v.length != 4).findAny().isPresent()) {
            throw new ExcelParseException("array of Integers can only be 4 sizes in mergeData ");
        }
        this.mergeData = mergeData;
    }
}
