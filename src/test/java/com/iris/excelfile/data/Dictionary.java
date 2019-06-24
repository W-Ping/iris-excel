package com.iris.excelfile.data;

public enum Dictionary {
    A(1, "a"),

    B(340, "你好");

    private Integer code;
    private String value;

    private Dictionary(Integer code, String value) {
        this.code = code;
        this.value = value;
    }

    public static String getDictionaryValue(Integer code) {
        for (Dictionary em : values()) {
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
