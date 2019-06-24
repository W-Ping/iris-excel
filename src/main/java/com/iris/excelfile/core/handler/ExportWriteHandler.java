package com.iris.excelfile.core.handler;


import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.core.builder.IWriteBuilder;
import com.iris.excelfile.core.builder.impl.WriteBuilderImpl;
import com.iris.excelfile.metadata.ExcelSheet;
import org.apache.commons.collections.CollectionUtils;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author liu_wp
 * @date Created in 2019/3/6 10:20
 * @see
 */
public class ExportWriteHandler {
    private IWriteBuilder iExcelBuilder;

    public ExportWriteHandler(InputStream tempInputStream, OutputStream outputStream, ExcelTypeEnum excelTypeEnum) {
        this.iExcelBuilder = new WriteBuilderImpl(tempInputStream, outputStream, excelTypeEnum);
    }

    /**
     * @param excelSheets
     * @param isQueue
     */
    public void exportExcelV2007(List<ExcelSheet> excelSheets, boolean isQueue) {
        if (!CollectionUtils.isEmpty(excelSheets)) {
            for (ExcelSheet excelSheet : excelSheets) {
                excelSheet.setQueueTask(isQueue);
                iExcelBuilder.addContentToExcel(excelSheet);
            }
            iExcelBuilder.finish();
        }
    }

}
