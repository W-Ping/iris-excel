package com.iris.excelfile.handler;

import com.iris.excelfile.core.handler.extend.IWriteLoadTemplateHandler;
import com.iris.excelfile.metadata.ExcelFont;
import com.iris.excelfile.utils.TemplateUtil;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 * @author liu_wp
 * @date Created in 2019/6/21 11:13
 * @see
 */
public class CrossWriteLoadTemplateHandlerImpl implements IWriteLoadTemplateHandler {
    /**
     *
     */
    private Map<String, String> refMap = new HashMap<>();

    public CrossWriteLoadTemplateHandlerImpl(Map<String, String> refMap) {
        this.refMap = refMap;
    }

    /**
     * @param workbook
     * @param sheet
     */
    @Override
    public void beforeLoadTemplate(Workbook workbook, Sheet sheet) {
        ExcelFont excelFont = new ExcelFont();
        excelFont.setFontColor(IndexedColors.RED);
        excelFont.setBold(true);
        excelFont.setFontHeightInPoints((short) 12);
        TemplateUtil.replaceTemplate(workbook, sheet, excelFont, 2, 1, -1, 2, refMap, "YYYY", "MM");
        TemplateUtil.replaceTemplate(workbook, sheet, excelFont, 2, 34, -1, 2, refMap, "YYYY", "MM");
        TemplateUtil.replaceTemplate(workbook, sheet, excelFont, 2, 47, -1, 2, refMap, "YYYY", "MM");
        TemplateUtil.replaceTemplate(workbook, sheet, excelFont, 13, 1, -1, 2, refMap, "YYYY", "MM");
        TemplateUtil.replaceTemplate(workbook, sheet, excelFont, 13, 34, -1, 2, refMap, "YYYY", "MM");

    }

    public Map<String, String> getRefMap() {
        return refMap;
    }

    public void setRefMap(Map<String, String> refMap) {
        this.refMap = refMap;
    }
}
