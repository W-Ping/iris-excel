package com.iris.excelfile.core.handler.extend;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author liu_wp
 * @date Created in 2019/6/21 11:07
 * @see
 */
public interface IWriteLoadTemplateHandler {

    /**
     * @param sheet
     */
    void beforeLoadTemplate(Workbook workbook, Sheet sheet);
}
