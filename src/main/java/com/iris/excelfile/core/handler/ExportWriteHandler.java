package com.iris.excelfile.core.handler;


import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.core.builder.IWriteBuilder;
import com.iris.excelfile.core.builder.impl.WriteBuilderImpl;
import com.iris.excelfile.exception.ExcelException;
import com.iris.excelfile.metadata.ExcelSheet;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 10:20
 * @see
 */
@Slf4j
public class ExportWriteHandler {
	private IWriteBuilder iWriteBuilder;

	public ExportWriteHandler(InputStream tempInputStream, OutputStream outputStream, String excelOutFileFullPath, ExcelTypeEnum excelTypeEnum) {
		this.iWriteBuilder = new WriteBuilderImpl(tempInputStream, outputStream, excelOutFileFullPath, excelTypeEnum);
	}

	/**
	 * @param excelSheets
	 * @param isQueue
	 * @throws ExcelException
	 */
	public void exportExcelV2007(List<ExcelSheet> excelSheets, boolean isQueue) throws ExcelException {
		if (!CollectionUtils.isEmpty(excelSheets)) {
			try {
				for (ExcelSheet excelSheet : excelSheets) {
					iWriteBuilder.addContentToExcel(excelSheet, isQueue);
				}
				iWriteBuilder.flush();
				iWriteBuilder.close();
			} catch (IOException e) {
				e.printStackTrace();
				log.error("数据导出IO异常{}", e.getMessage() != null ? e.getMessage() : e.toString());
				throw new ExcelException("数据导出IO异常", e);
			}
		}
	}
}
