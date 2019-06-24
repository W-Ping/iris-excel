package com.iris.excelfile.data;

public enum Dictionary2 {
    A(8, "发发发"),

    B(340, "辅导费");

    private Integer code;
    private String value;

    private Dictionary2(Integer code, String value) {
        this.code = code;
        this.value = value;
    }

    public static String getDictionaryValue(Integer code) {
        for (Dictionary2 em : values()) {
            if (em.getCode().equals(code)) {
                return em.getValue();
            }
        }
        return null;
    }

    public Integer getCode() {
        return code;
    }

    public void setCode(Integer code) {
        this.code = code;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}
