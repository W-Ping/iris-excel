package com.iris.excelfile.core.builder.impl;

import com.iris.excelfile.constant.BorderEnum;
import com.iris.excelfile.constant.DataFormatEnum;
import com.iris.excelfile.constant.ExcelFormula;
import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.context.WriteContext;
import com.iris.excelfile.core.builder.AbstractWriteBuilder;
import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.*;
import com.iris.excelfile.utils.JSONUtil;
import com.iris.excelfile.utils.StyleUtil;
import com.iris.excelfile.utils.TypeUtil;
import com.iris.excelfile.utils.WorkBookUtil;
import lombok.extern.slf4j.Slf4j;
import net.sf.cglib.beans.BeanMap;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 10:22
 * @see
 */
@Slf4j
public class WriteBuilderImpl extends AbstractWriteBuilder {

    private WriteContext context;
    /**
     * 保存已经合并的行索引值
     */
    private Map<Integer, Integer> mergeRowIndexMap = new HashMap<>();


    public WriteBuilderImpl(InputStream templateInputStream,
                            OutputStream outputStream, String excelOutFileFullPath,
                            ExcelTypeEnum excelType) {
        try {
            context = new WriteContext(templateInputStream, outputStream, excelOutFileFullPath, excelType);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * @param excelSheet
     */
    @Override
    public void initSheet(ExcelSheet excelSheet) {
        this.clearCache();
        context.currentSheet(excelSheet);
    }

    /**
     * @param excelTable
     */
    @Override
    public void initTable(ExcelTable excelTable) {
        if (excelTable == null) {
            return;
        }
        context.initTable(excelTable);
    }

    /**
     * @param table
     */
    @Override
    public void addContentBefore(ExcelTable table) {
        this.initDataContent(table);
    }

    /**
     * @param table
     */
    @Override
    public boolean addContent(ExcelTable table) {
        if (table == null) {
            return false;
        }
        try {
            List<T> data = table.getData();
            for (int i = 0; i < data.size(); i++) {
                addOneRowDataToExcel(data.get(i), table, i);
//            StyleUtil.buildCellBorderStyle(currentCellStyle, 0, i == 0 && !table.isNeedHead(), BorderEnum.TOP);
            }
            return true;
        } catch (Exception e) {
            log.error("excel 加载数据失败 {}", e.getMessage());
            throw new ExcelParseException("excel 加载数据失败", e);
        }
    }


    /**
     * @param startRow
     * @param endRow
     * @param startCell
     * @param endCell
     */
    @Override
    public void merge(int startRow, int endRow, int startCell, int endCell) {
        context.merge(startRow, endRow, startCell, endCell);
    }


    /**
     * @param table
     */
    private void initDataContent(ExcelTable table) {
        int startRow = table.getStartContentRowIndex();
        log.info("table{}({}) execute addContent method startContentRowIndex【{}】", table.getTableNo(), table.getTableName(), startRow);
        ExcelStyle tableStyle = table.getTableStyle();
        CellStyle currentCellStyle = null;
        if (tableStyle != null) {
            currentCellStyle = tableStyle.getCurrentCellStyle();
            StyleUtil.buildAllCellDataFormatStyle(context.getWorkbook(), currentCellStyle, table.getTableDataDefaultFormat(), true);
        }
        Map<Integer, CellStyle> cellStyleMap = null;
        //加载数据字典和自定义样式
        if (table.getWriteBeforeHandler() != null) {
            cellStyleMap = table.getWriteBeforeHandler().customSetCellStyle(context.getWorkbook(), DataFormatEnum.getDataFormatShort(context.getWorkbook(), table.getTableDataDefaultFormat()));
            table.setCellStyleMap(cellStyleMap);
        }
        if (context.isNeedInputTemplate()) {
            Sheet currentSheet = context.getCurrentSheet();
            ExcelShiftRange excelShiftRange = table.getExcelShiftRange();
            if (null != excelShiftRange) {
                startRow = excelShiftRange.getStartRow();
                log.info("table{},excelShiftRange:【{}】", table.getTableNo(), JSONUtil.objectToString(excelShiftRange));
                int startRow1 = excelShiftRange.getStartRow();
                int shiftNumber = excelShiftRange.getShiftNumber();
                int i = excelShiftRange.getEndRow() > 0
                        && excelShiftRange.getEndRow() > excelShiftRange.getStartRow() ? excelShiftRange.getEndRow() : excelShiftRange.getStartRow() > currentSheet.getLastRowNum() ? excelShiftRange.getStartRow() : currentSheet.getLastRowNum();
                try {
                    currentSheet.shiftRows(startRow1, i,
                            shiftNumber, excelShiftRange.isCopyRowHeight(), excelShiftRange.isResetOriginalRowHeight());
                } catch (Exception e) {
                    log.error("excel shiftRows fail {}", e.getMessage());
                    throw new ExcelParseException("excel shiftRows fail", e);
                }
            }
        }
        table.setStartContentRowIndex(startRow);
        tableStyle.setCurrentCellStyle(currentCellStyle);
        table.setTableStyle(tableStyle);
    }

    /**
     * @param oneRowData
     * @param excelTable
     * @param i
     */
    public void addOneRowDataToExcel(Object oneRowData, ExcelTable excelTable, int i) {
        int n = excelTable.getStartContentRowIndex() + i;
        CellStyle cellStyle = excelTable.getTableStyle().getCurrentCellStyle();
        Row row = WorkBookUtil.createOrGetRow(context.getCurrentSheet(), n);
        if (excelTable.getHeightInPoints() != null) {
            row.setHeightInPoints(excelTable.getHeightInPoints());
        }
        if (excelTable.getWriteAfterHandler() != null) {
            excelTable.getWriteAfterHandler().row(n, row);
        }
        if (oneRowData instanceof List) {
            addListTypeToExcel((List) oneRowData, row, cellStyle, excelTable, i);
        } else {
            addObjectTypeToExcel(oneRowData, row, cellStyle, excelTable, i);
        }

    }

    /**
     * @param oneRowData
     * @param row
     * @param cellStyle
     * @param excelTable
     * @param loop
     */
    private void addListTypeToExcel(List<Object> oneRowData, Row row, CellStyle cellStyle, ExcelTable excelTable, int loop) {
        if (CollectionUtils.isEmpty(oneRowData)) {
            return;
        }
        for (int i = 0; i < oneRowData.size(); i++) {
            Object cellValue = oneRowData.get(i);
            Cell cell = WorkBookUtil.createCell(row, i + excelTable.getFirstCellIndex(), cellStyle, cellValue,
                    TypeUtil.isNum(cellValue), false, false);
            if (excelTable.getWriteAfterHandler() != null) {
                excelTable.getWriteAfterHandler().cell(i, cell);
            }
        }
    }

    /**
     * @param oneRowData
     * @param row
     * @param cellStyle
     * @param excelTable
     * @param loop
     */
    private void addObjectTypeToExcel(Object oneRowData, Row row, CellStyle cellStyle, ExcelTable excelTable, int loop) {
        int startCellIndex = excelTable.getFirstCellIndex();
        BeanMap beanMap = BeanMap.create(oneRowData);
        List<ExcelColumnProperty> columnPropertyList = excelTable.getExcelHeadProperty().getColumnPropertyList();
        if (!CollectionUtils.isEmpty(columnPropertyList)) {
            //是否为计算公式
            boolean isFormula = false;
            //是否为空列
            boolean isNullField = false;
            List<String> sumCellFormula = null;
            String divideCellFormula = null;
            Integer mergeCellIndex = null;
            int mergeRowCount = 0;
            String dateFormat = null;
            String cellValue = null;
            CellStyle fieldIsNullStyle = null;
            boolean isSeqNo = false;
            int lastIndex = columnPropertyList.size() - 1;
            for (int i = 0; i < columnPropertyList.size(); i++) {
                ExcelColumnProperty excelColumnProperty = columnPropertyList.get(i);
                if (excelColumnProperty == null || excelColumnProperty.isIgnoreCell()) {
                    continue;
                }
                StyleUtil.buildCellBorderStyle(cellStyle, 0, 0 == i, BorderEnum.LEFT);
                StyleUtil.buildCellBorderStyle(cellStyle, lastIndex, lastIndex == i, BorderEnum.RIGHT);
                CellStyle cellStyle1 = cellStyle;
                Map<Integer, CellStyle> cellStyleMap = excelTable.getCellStyleMap();
                if (null != cellStyleMap && null != cellStyleMap.get(i + startCellIndex)) {
                    cellStyle1 = cellStyleMap.get(i + startCellIndex);
                }
                isSeqNo = excelColumnProperty.isSeqNo();
                String dicValue = null;
                //合并行和合并列
                mergeCellIndex = excelColumnProperty.getMergeCellIndex();
                mergeRowCount = excelColumnProperty.getMergeRowCount();
                if (mergeRowCount != 0 || mergeCellIndex != null) {
                    if (mergeCellIndex != null && mergeRowCount == 0) {
                        merge(row.getRowNum(), row.getRowNum(), i + startCellIndex, mergeCellIndex);
                    } else if (mergeRowCount != 0) {
                        Integer mergeMaxRowIndex = mergeRowIndexMap.get(excelTable.getTableNo());
                        int endRowIndex = row.getRowNum() + mergeRowCount;
                        if (mergeMaxRowIndex == null) {
                            //第一次合并
                            merge(row.getRowNum(), endRowIndex, i + startCellIndex, mergeCellIndex != null ? mergeCellIndex : i + startCellIndex);
                            mergeRowIndexMap.put(excelTable.getTableNo(), endRowIndex);
                        } else {
                            if (mergeMaxRowIndex < row.getRowNum()) {
                                //合并
                                merge(row.getRowNum(), endRowIndex, i + startCellIndex, mergeCellIndex != null ? mergeCellIndex : i + startCellIndex);
                                mergeRowIndexMap.put(excelTable.getTableNo(), endRowIndex);
                            }
                        }
                    }
                }
                if (isSeqNo) {
                    int seqNumValue = (loop + 1);
                    cellValue = String.valueOf(seqNumValue);
                } else {
                    isNullField = excelColumnProperty.getField() == null;
                    sumCellFormula = excelColumnProperty.getSumCellFormula();
                    divideCellFormula = excelColumnProperty.getDivideCellFormula();
                    dateFormat = excelColumnProperty.getDateFormat();
                    //使用SUM和divide计算公式
                    isFormula = !CollectionUtils.isEmpty(sumCellFormula) || StringUtils.isNotBlank(divideCellFormula);
                    if (excelTable.isEffectExcelFormula() && isFormula) {
                        if (!CollectionUtils.isEmpty(sumCellFormula)) {
                            cellValue = ExcelFormula.assemblyExcelFormula(row.getRowNum() + 1, sumCellFormula, ExcelFormula.ADDITION);
                        } else if (StringUtils.isNotBlank(divideCellFormula)) {
                            Integer divideRow = row.getRowNum() + 1;
                            ExcelTable refTable = context.getCurrentExcelSheetParam().getSheetTablesCache().get(excelTable.getDivideFormulaTRefTable());
                            if (refTable != null) {
                                divideRow = refTable.getStartContentRowIndex() + 1;
                            }
                            cellValue = ExcelFormula.assemblyExcelFormula(divideRow + loop, divideCellFormula, ExcelFormula.DIVISION);
                        }
                    } else {
                        //空列或手动清空列
                        if (isNullField ||
                                excelTable.getWriteBeforeHandler() != null && excelTable.getWriteBeforeHandler().clearCellValue(row, i + startCellIndex)) {
                            cellValue = null;
                        } else {
                            cellValue = TypeUtil.getFieldStringValue(beanMap, excelColumnProperty.getField().getName(),
                                    dateFormat);
                        }
                        //数据字典转换 默认原值
                        if (StringUtils.isNotBlank(cellValue) && TypeUtil.isIntNum(excelColumnProperty.getField()) && excelTable.getDictionaryRefHandler() != null) {
                            dicValue = getDictValue(i, Integer.valueOf(cellValue), excelTable.getDictionaryRefHandler());
                            if (StringUtils.isNotBlank(dicValue)) {
                                cellValue = dicValue;
                            }
                        }
                    }
                    //空格样式 没有值
                    fieldIsNullStyle = StyleUtil.buildFieldIsNullStyle(context.getWorkbook(), isNullField);
                    if (fieldIsNullStyle != null) {
                        cellStyle1 = fieldIsNullStyle;
                    }
                }
                Cell cell = WorkBookUtil.createCell(row, i + startCellIndex, cellStyle1, cellValue,
                        cellValue != null && !isFormula && StringUtils.isBlank(dicValue)
                                ? TypeUtil.isNum(excelColumnProperty.getField()) : false, isFormula, isSeqNo);
                if (excelTable.getWriteAfterHandler() != null) {
                    excelTable.getWriteAfterHandler().cell(i, cell);
                }
            }
        }
    }

    private static String getDictValue(Integer str1, Integer str2, BiFunction<Integer, Integer, String> function) {
        return function.apply(str1, str2);
    }

    @Override
    public void clearCache() {
        //清空默认样式缓存
        StyleUtil.clearDefaultStyleCache();
        //TODO 多线程导出有跨行要修改
        mergeRowIndexMap.clear();
    }

    @Override
    public void flush() throws IOException {
//        context.getWorkbook().getCreationHelper().createFormulaEvaluator().evaluateAll();
        context.getWorkbook().write(context.getOutputStream());
    }

    @Override
    public void close() throws IOException {
        context.getWorkbook().close();
    }

    @Override
    public String getOutFilePath() {
        return context.getExcelOutFileFullPath();
    }
}
