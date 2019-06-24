package com.iris.excelfile.constant;

/**
 * @author liu_wp
 * @date Created in 2019/3/25 12:04
 * @see
 */
public enum DateTypeEnum {
    TIMESTAMP_TIME("yyyy-MM-dd HH:mm:ss"),
    TIMESTAMP_ZH_TIME("yyyy年MM月dd日"),
    TIMESTAMP_OBLIQUE_TIME("yyyy/MM/dd HH:mm:ss"),
    DATE("YYYY-MM-dd"),
    DATE_DEFAULT("dd/MM/yy"),
    DATE_OBLIQUE("YYYY/MM/dd"),
    TIME("HH:mm:ss");

    private final String format;

    private DateTypeEnum(String format) {
        this.format = format;
    }

    public String getFormat() {
        return format;
    }


}
