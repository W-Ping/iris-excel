package com.iris.excelfile;

import com.iris.excelfile.model.ReadExcelModel1;
import com.iris.excelfile.request.ExcelReadParam;
import com.iris.excelfile.response.ExcelReadResponse;
import com.iris.excelfile.utils.JSONUtil;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * @author liu_wp
 * @date 2019/9/6 15:27
 * @see
 */
public class ExcelReadTest {
    @Test
    public void readOneSheetExcelV2007() {
        ExcelReadParam excelReadParam = new ExcelReadParam();
        String excelPath = "D:\\lwp-work\\iris-excel\\src\\test\\resources\\readExcel1.xlsx";
        excelReadParam.setExcelPath(excelPath);
        List<Class> classes = new ArrayList<>();
        classes.add(ReadExcelModel1.class);
        excelReadParam.setDataClass(classes);
        ExcelReadResponse excelReadResponse = IRISExcelFileFactory.readOneSheetExcelV2007(excelReadParam);
        System.out.println(JSONUtil.objectToString(excelReadResponse));
    }
}
