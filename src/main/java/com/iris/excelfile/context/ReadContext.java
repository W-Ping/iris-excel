package com.iris.excelfile.context;

import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.utils.WorkBookUtil;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;

/**
 * @author liu_wp
 * @date 2019/8/30 9:23
 * @see
 */
public class ReadContext {
    private Workbook workbook;

    public ReadContext(InputStream inputStream, ExcelTypeEnum excelType) {
        try {
            this.workbook = WorkBookUtil.createWorkBook(inputStream, excelType);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }
}
