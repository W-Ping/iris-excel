package com.iris.excelfile.core.handler;

import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.core.builder.IReadBuilder;
import com.iris.excelfile.core.builder.impl.ReadBuiderImpl;
import org.apache.poi.ss.formula.functions.T;

import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * @author liu_wp
 * @date 2019/8/30 9:47
 * @see
 */
public class ExcelReadHandler {
    private IReadBuilder iReadBuilder;

    public ExcelReadHandler(InputStream inputStream, ExcelTypeEnum excelType) {
        iReadBuilder = new ReadBuiderImpl(inputStream, excelType);
    }

    public Map<Integer, Object> readExcelV2007(List<Class> classes) {
        iReadBuilder.initDataClass(classes);
        return iReadBuilder.addExcelToObject();
    }

    public Object readOneSheetExcelV2007(Class classes) {
        iReadBuilder.initDataClass(Arrays.asList(classes));
        return iReadBuilder.addOneSheetExcelToObject();
    }
}
