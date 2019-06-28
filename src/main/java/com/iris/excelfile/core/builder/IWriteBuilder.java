package com.iris.excelfile.core.builder;

import com.iris.excelfile.exception.ExcelException;
import com.iris.excelfile.metadata.ExcelSheet;
import com.iris.excelfile.metadata.ExcelTable;

import java.io.IOException;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 10:14
 * @see
 */
public interface IWriteBuilder {


    /**
     * 导出文档入口
     *
     * @param excelSheet
     */
    void addContentToExcel(ExcelSheet excelSheet) throws ExcelException;

    /**
     * 初始化sheet
     *
     * @param excelSheet
     */
    void initSheet(ExcelSheet excelSheet);

    /**
     * 初始化sheet 之后
     *
     * @param excelSheet
     */
    void initSheetAfter(ExcelSheet excelSheet) throws ExcelException;


    /**
     * 初始化 table
     *
     * @param excelSheet
     */
    void initTable(ExcelTable excelTable);

    /**
     * 加载数据之前处理
     *
     * @param table
     */
    void addContentBefore(ExcelTable table);

    /**
     * 加载数据
     *
     * @param table
     */
    boolean addContent(ExcelTable table);


    /**
     * 合并行列
     *
     * @param startRow
     * @param endRow
     * @param startCell
     * @param endCell
     */
    void merge(int startRow, int endRow, int startCell, int endCell);

    /**
     * 清除缓存
     */
    void clearCache();

    /**
     * 写出文档
     */
    void finish() throws IOException;


    /**
     * 导出文件路径
     *
     * @return
     */
    String getOutFilePath();
}
