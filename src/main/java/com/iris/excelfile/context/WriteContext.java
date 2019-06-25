package com.iris.excelfile.context;


import com.iris.excelfile.constant.BorderEnum;
import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.constant.FileConstant;
import com.iris.excelfile.constant.TableBodyEnum;
import com.iris.excelfile.metadata.*;
import com.iris.excelfile.utils.StyleUtil;
import com.iris.excelfile.utils.WorkBookUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 10:00
 * @see
 */
public class WriteContext {
    private Workbook workbook;
    private Sheet currentSheet;
    private String currentSheetName;
    private ExcelSheet currentExcelSheetParam;
    private ExcelTypeEnum excelType;
    private OutputStream outputStream;
    private boolean needInputTemplate;
    private String excelOutFileFullPath;

    public WriteContext(InputStream templateInputStream, OutputStream outputStream, String excelOutFileFullPath, ExcelTypeEnum excelType) throws IOException {
        this.excelType = excelType;
        this.outputStream = outputStream;
        this.needInputTemplate = templateInputStream != null;
        this.workbook = WorkBookUtil.createWorkBook(templateInputStream, excelType);
    }

    public void currentSheet(ExcelSheet excelSheet) {
        if (null == currentExcelSheetParam || currentExcelSheetParam.getSheetNo() != excelSheet.getSheetNo()) {
            cleanCurrentSheet();
            this.currentExcelSheetParam = excelSheet;
            this.currentSheetName = excelSheet.getSheetName();
            try {
                this.currentSheet = this.workbook.getSheetAt(excelSheet.getSheetNo());
            } catch (Exception e) {
                this.currentSheet = WorkBookUtil.createSheet(this.workbook, excelSheet);
            }
            this.initSheet(excelSheet);
            excelSheet.initCurrentSheetTables(excelSheet);
        }

    }

    /**
     * @param excelSheet
     */
    public void initSheet(ExcelSheet excelSheet) {
        //加载处理模板
        if (this.needInputTemplate && excelSheet.getWriteLoadTemplateHandler() != null) {
            excelSheet.getWriteLoadTemplateHandler().beforeLoadTemplate(this.workbook, this.currentSheet);
        }
        //合并单元格
        if (excelSheet.getMergeData() != null) {
            excelSheet.getMergeData().stream().forEach(v -> this.merge(v[0], v[1], v[2], v[3]));
        }
        //冻结单元格
        ExcelFreezePaneRange freezePaneRange = excelSheet.getFreezePaneRange();
        if (freezePaneRange != null) {
            this.currentSheet.createFreezePane(freezePaneRange.getCellCount(), freezePaneRange.getRowCount(), freezePaneRange.getCellIndex(), freezePaneRange.getRowIndex());
        }
        //设置保护密码
        this.currentSheet.protectSheet(StringUtils.isNotBlank(excelSheet.getProtectSheetPwd()) ? excelSheet.getProtectSheetPwd() : FileConstant.PROTECT_SHEET_PASSWORD);
        //设置列宽
        StyleUtil.buildTableWidthStyle(this.getCurrentSheet(), this.currentExcelSheetParam.getColumnWidthMap());
    }

    public void merge(int startRow, int endRow, int startCell, int endCell) {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(startRow, endRow, startCell, endCell);
        this.currentSheet.addMergedRegion(cellRangeAddress);
    }

    /**
     * init table
     *
     * @param table
     */
    public void initTable(ExcelTable table) {
        initTableHeadStyle(table);
        initTableContentStyle(table);
        initTableHead(table);
    }

    private void initTableHead(ExcelTable table) {
        if (!table.isNeedHead()) {
            return;
        }
        ExcelHeadProperty excelHeadProperty = table.getExcelHeadProperty();
        ExcelStyle headStyle = excelHeadProperty.getHeadStyle();
        if (excelHeadProperty != null && excelHeadProperty.getHead().size() > 0) {
            int startRow = table.getFirstRowIndex();
            try {
                addMergedRegionToCurrentSheet(startRow, table.getFirstCellIndex(), table.getExcelHeadProperty());
            } catch (Exception e) {
                e.printStackTrace();
            }
            int rowNum = excelHeadProperty.getRowNum();
            for (int i = startRow; i < rowNum + startRow; i++) {
                Row row = WorkBookUtil.createOrGetRow(currentSheet, i);
                if (headStyle != null) {
                    StyleUtil.buildCellBorderStyle(headStyle.getCurrentCellStyle(), startRow, i == startRow, BorderEnum.TOP);
                }
                addOneRowOfHeadDataToExcel(row, excelHeadProperty.getHeadByRowNum(i - startRow), table.getFirstCellIndex(), headStyle);
            }
        }
    }

    /**
     * table 头样式
     *
     * @param table
     */
    private void initTableHeadStyle(ExcelTable table) {
        ExcelHeadProperty excelHeadProperty = table.getExcelHeadProperty();
        ExcelStyle headStyle = excelHeadProperty.getHeadStyle();
        CellStyle cellStyle = null;
        if (headStyle != null) {
            cellStyle = StyleUtil.buildCellStyle(this.workbook, table.getTableNo(), headStyle.getExcelFont(),
                    headStyle.getExcelBackGroundColor(), TableBodyEnum.HEAD, table.isLocked());
        } else {
            headStyle = new ExcelStyle();
            cellStyle = StyleUtil.defaultHeadStyle(this.workbook, table.getTableNo(), table.isLocked());
        }
        headStyle.setCurrentCellStyle(cellStyle);
        table.setTableHeadStyle(headStyle);
    }

    /**
     * table 内容样式
     *
     * @param table
     */
    private void initTableContentStyle(ExcelTable table) {
        ExcelStyle tableStyle = table.getTableStyle();
        CellStyle cellStyle = null;
        if (tableStyle != null) {
            cellStyle = StyleUtil.buildCellStyle(this.workbook, table.getTableNo(), tableStyle.getExcelFont(),
                    tableStyle.getExcelBackGroundColor(), TableBodyEnum.CONTENT, table.isLocked());
            tableStyle.setCurrentCellStyle(cellStyle);
        } else {
            tableStyle = new ExcelStyle();
            cellStyle = StyleUtil.defaultContentCellStyle(this.workbook, table.getTableNo(), table.isLocked());
            tableStyle.setDefaultCellStyle(cellStyle);
        }
        tableStyle.setCurrentCellStyle(cellStyle);
        table.setTableStyle(tableStyle);
    }

    /**
     * @param row
     * @param headByRowNum
     * @param startCellIndex
     * @param tableStyle
     */
    private void addOneRowOfHeadDataToExcel(Row row, List<String> headByRowNum, int startCellIndex, ExcelStyle tableStyle) {
        CellStyle currentCellStyle = null;
        if (tableStyle != null) {
            currentCellStyle = tableStyle.getCurrentCellStyle();
        }
        if (headByRowNum != null && headByRowNum.size() > 0) {
            int lastIndex = headByRowNum.size() - 1;
            for (int i = 0; i < headByRowNum.size(); i++) {
                StyleUtil.buildCellBorderStyle(currentCellStyle, 0, 0 == i, BorderEnum.LEFT);
                StyleUtil.buildCellBorderStyle(currentCellStyle, lastIndex, lastIndex == i, BorderEnum.RIGHT);
                WorkBookUtil.createCell(row, startCellIndex + i, currentCellStyle, headByRowNum.get(i));
            }

        }
    }

    private void addMergedRegionToCurrentSheet(int startRow, int startCell, ExcelHeadProperty excelHeadProperty) {
        for (ExcelCellRange cellRangeModel : excelHeadProperty.getCellRangeModels(startCell)) {
            currentSheet.addMergedRegion(new CellRangeAddress(cellRangeModel.getFirstRowIndex() + startRow,
                    cellRangeModel.getLastRowIndex() + startRow,
                    cellRangeModel.getFirstCellIndex(), cellRangeModel.getLastCellIndex()));

        }
    }

    private void cleanCurrentSheet() {
        this.currentSheet = null;
        this.currentSheetName = null;
        this.currentExcelSheetParam = null;

    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public String getCurrentSheetName() {
        return currentSheetName;
    }

    public void setCurrentSheetName(String currentSheetName) {
        this.currentSheetName = currentSheetName;
    }

    public ExcelSheet getCurrentExcelSheetParam() {
        return currentExcelSheetParam;
    }

    public void setCurrentExcelSheetParam(ExcelSheet currentExcelSheetParam) {
        this.currentExcelSheetParam = currentExcelSheetParam;
    }

    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    /**
     * @return
     */
    public boolean isNeedInputTemplate() {
        return needInputTemplate;
    }

    public void setNeedInputTemplate(boolean needInputTemplate) {
        this.needInputTemplate = needInputTemplate;
    }

    public String getExcelOutFileFullPath() {
        return excelOutFileFullPath;
    }

    public void setExcelOutFileFullPath(String excelOutFileFullPath) {
        this.excelOutFileFullPath = excelOutFileFullPath;
    }
}
