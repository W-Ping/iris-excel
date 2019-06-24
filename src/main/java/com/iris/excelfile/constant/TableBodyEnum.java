package com.iris.excelfile.constant;

/**
 * @author liu_wp
 * @date Created in 2019/3/8 11:19
 * @see
 */
public enum TableBodyEnum {
    HEAD("head"), CONTENT("content"), FOOT("foot");
    private String name;

    private TableBodyEnum(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }
}
