package com.iris.excelfile.core.builder;

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
	 * @param queueTask
	 */
	void addContentToExcel(ExcelSheet excelSheet, boolean queueTask);

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
	 * @param queueTask
	 */
	void initSheetAfter(ExcelSheet excelSheet, boolean queueTask);


	/**
	 * 初始化 table
	 *
	 * @param excelTable
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
	 * @return
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
	 * 写入文档
	 *
	 * @throws IOException
	 */
	void flush() throws IOException;

	/**
	 * 关闭导出流
	 *
	 * @throws IOException
	 */
	void close() throws IOException;

	/**
	 * 获取导出文件路径
	 *
	 * @return
	 */
	String getOutFilePath();
}
