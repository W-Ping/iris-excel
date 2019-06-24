package com.iris.excelfile.constant;


import java.io.IOException;
import java.io.InputStream;

/**
 * @author liu_wp
 * @date Created in 2019/3/1 10:19
 * @see
 */
public enum ExcelTypeEnum {

    XLS(".xls"), XLSX(".xlsx");

    private String value;

    ExcelTypeEnum(String value) {
        this.setValue(value);
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

}
