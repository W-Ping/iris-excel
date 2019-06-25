package com.iris.excelfile.core.builder;

import com.iris.excelfile.exception.ExcelException;
import com.iris.excelfile.metadata.ExcelSheet;
import com.iris.excelfile.metadata.ExcelTable;
import com.iris.excelfile.queue.multitask.ExcelWriteQueueTaskHandler;

import java.util.List;

/**
 * @author liu_wp
 * @date Created in 2019/6/24 10:24
 * @see
 */
public abstract class AbstractWriteBuilder implements IWriteBuilder {
    protected ExcelWriteQueueTaskHandler excelWriteQueueServiceImpl;

    /**
     * @param excelSheet
     */
    @Override
    public void addContentToExcel(ExcelSheet excelSheet) throws ExcelException {
        initSheet(excelSheet);
        initSheetAfter(excelSheet);
    }

    /**
     * 加载数据之前处理
     *
     * @param excelSheet
     */
    @Override
    public void initSheetAfter(ExcelSheet excelSheet) throws ExcelException {
        List<ExcelTable> excelTables = excelSheet.getExcelTables();
        boolean queueTask = excelSheet.isQueueTask();
        //初始化队列
        if (queueTask) {
            excelWriteQueueServiceImpl = new ExcelWriteQueueTaskHandler(excelTables.size(), this);
        }
        for (ExcelTable excelTable : excelTables) {
            initTable(excelTable);
            addContentBefore(excelTable);
            //加入队列
            if (queueTask) {
                excelWriteQueueServiceImpl.putQueue(excelTable);
            } else {
                addContent(excelTable);
            }
        }
        //多线程处理
        if (queueTask) {
            excelWriteQueueServiceImpl.consumeQueue();
        }
    }
}
