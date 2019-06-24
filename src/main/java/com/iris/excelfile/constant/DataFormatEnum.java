package com.iris.excelfile.constant;

import com.iris.excelfile.exception.ExcelParseException;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 数据格式枚举
 *
 * @author liu_wp
 * @date Created in 2019/6/19 10:28
 * @see
 */
public enum DataFormatEnum {
    NORMAL(""),
    NUMBER_1_CURRENCY_CN("￥#,##0"),
    NUMBER_1_CURRENCY_US("$#,##0.00"),
    NUMBER_1_KILOBIT("#,##0"),
    NUMBER_1_PERCENT("0%"),
    NUMBER_1_ROUND_UP("0"),
    NUMBER_2_CURRENCY_CN("￥#,##0.00"),
    NUMBER_2_CURRENCY_US("$#,##0.00"),
    NUMBER_2_KILOBIT("#,##0.00"),
    NUMBER_2_PERCENT("0.00%"),
    NUMBER_2_ROUND_UP("0.00"),
    NUMBER_4_CURRENCY_US("$#,##0.0000"),
    NUMBER_4_KILOBIT("#,##0.0000"),
    NUMBER_4_PERCENT("0.0000%"),
    NUMBER_4_ROUND_UP("0.0000"),
    NUMBER_F_E_1("0.00E+00"),
    NUMBER_F_E_2("##0.0E+0"),
    /**
     * 数字转大写中文
     */
    NUMBER_UPPERCASE_CN_LANGUAGE("[DbNum2][$-804]0");

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    private String format;

    public static short getDataFormatShort(Workbook workbook, DataFormatEnum dataFormatEnum) {
        if (dataFormatEnum == null) {
            return 0;
        }
        boolean isOk = false;
        for (DataFormatEnum em : values()) {
            if (dataFormatEnum == em) {
                isOk = true;
                break;
            }
        }
        if (!isOk) {
            throw new ExcelParseException(dataFormatEnum + " dataFormat is Undefined");
        }
        DataFormat dataFormat = workbook.createDataFormat();
        return dataFormat.getFormat(dataFormatEnum.getFormat());
    }

    private DataFormatEnum(String format) {
        this.format = format;
    }
}
