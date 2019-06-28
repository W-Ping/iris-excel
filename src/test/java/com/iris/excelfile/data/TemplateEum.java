package com.iris.excelfile.data;

/**
 * @author liu_wp
 * @date Created in 2019/6/28 17:36
 * @see
 */
public enum TemplateEum {
    ADMINISTRATOR(1, "管理员"),
    GENERAL(2, "普通用户");
    private Integer type;
    private String name;

    private TemplateEum(Integer type, String name) {
        this.type = type;
        this.name = name;
    }

    public static String getTypeName(Integer type) {
        if (type == null) {
            return null;
        }
        for (TemplateEum value : values()) {
            if (value.getType() == type) {
                return value.getName();
            }
        }
        return null;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
