package com.iris.excelfile.core.builder;

import com.iris.excelfile.constant.FileConstant;
import com.iris.excelfile.metadata.ExcelSheet;
import com.iris.excelfile.metadata.ExcelTable;
import com.iris.excelfile.queue.IExcelWriteQueueTaskHandler;
import com.iris.excelfile.queue.multitask.ExcelWriteQueueTaskHandler;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

/**
 * @author liu_wp
 * @date Created in 2019/6/24 10:24
 * @see
 */
@Slf4j
public abstract class AbstractWriteBuilder implements IWriteBuilder {
	/**
	 *
	 */
	protected ExcelWriteQueueTaskHandler excelWriteQueueTaskHandler;


	/**
	 * @param excelSheet
	 * @param queueTask
	 */
	@Override
	public void addContentToExcel(ExcelSheet excelSheet, boolean queueTask) {
		initSheet(excelSheet);
		initSheetAfter(excelSheet, queueTask);
	}

	/**
	 * @param excelSheet
	 * @param queueTask
	 */
	@Override
	public void initSheetAfter(ExcelSheet excelSheet, boolean queueTask) {
		List<ExcelTable> excelTables = excelSheet.getExcelTables();
		queueTask = isQueueTaskRule(excelTables, queueTask);
		for (ExcelTable excelTable : excelTables) {
			initTable(excelTable);
			addContentBefore(excelTable);
			if (!queueTask) {
				addContent(excelTable);
			}
		}
		//多线程处理
		if (queueTask) {
			excelWriteQueueTaskHandler = new ExcelWriteQueueTaskHandler(excelTables.size(), this);
			//加入队列
			excelWriteQueueTaskHandler.putQueue(excelTables);
			//消费处理队列
			excelWriteQueueTaskHandler.consumeQueue();
		}
	}


	/**
	 * 判断是否加入队列处理规则
	 * 1：手动设置是否调用
	 * 2：数据不小于设置最小阈值(必要条件)
	 * 3：消息队列不小于设置最小阈值(必要条件)
	 *
	 * @param excelTables
	 * @param queueTask
	 * @return
	 */
	private boolean isQueueTaskRule(List<ExcelTable> excelTables, boolean queueTask) {
		int tableDataSizeTotal = excelTables.stream().mapToInt(v -> v.getData().size()).sum();
		if (queueTask) {
			queueTask = FileConstant.QUEUE_MIN_DATA_SIZE < tableDataSizeTotal;
		}
		return excelTables.size() >= FileConstant.QUEUE_MIN_SIZE ? queueTask : false;
	}
}
