package com.iris.excelfile.constant;

import com.iris.excelfile.exception.ExcelParseException;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author liu_wp
 * @date Created in 2019/6/20 13:23
 * @see
 */
@Slf4j
public enum ExcelFormula {
    SUBTRACTION("-"),
    ADDITION("+"),
    DIVISION("/"),
    MULTIPLICATION("*");
    private String type;

    private ExcelFormula(String type) {
        this.type = type;
    }

    public String assemblyFormula(String[] values) {
        boolean f = false;
        String collect = "";
        for (int i = 0; i < values.length; i++) {
            String value = values[i];
            f = Arrays.stream(ExcelFormula.values()).anyMatch(v -> value.indexOf(v.type) >= 0);
            if (f) {
                break;
            }
        }
        if (f) {
            collect = Arrays.stream(values).collect(Collectors.joining(this.type, "(", ")"));
        } else {
            collect = Arrays.stream(values).collect(Collectors.joining(":" + this.type, "(", ":)"));
        }
        return collect;
    }

    public static String assemblyExcelFormula(int row, String value) {
        if (log.isDebugEnabled()) {
            log.info("excelfile formula 【{}】", value);
        }
        for (ExcelFormula em : values()) {
            String type = em.type;
            if (value.indexOf(type) != -1) {
                String[] split = value.split("[" + type + "]");
                value = em.assemblyFormula(split);
                value = value.replaceAll("[:][" + type + "]", row + type);
            }
        }
        value = value.replaceAll("[:]", row + "");
        if (log.isDebugEnabled()) {
            log.info("excelfile formula value 【{}】", value);
        }
        return "IFERROR(" + value + ",\"\")";
    }

    public static String assemblyExcelFormula(int row, Object object, ExcelFormula excelFormula) {
        String value = "";
        if (object != null) {
            if (ExcelFormula.ADDITION.equals(excelFormula)) {
                List<String> list = (List<String>) object;
                //拼接格式：SUM(AD8,Z8,V8,R8,N8,J8,F8,B8)
                value = list.stream().collect(Collectors.joining(row + ",", "SUM(", row + ")"));
            } else if (ExcelFormula.DIVISION.equals(excelFormula)) {
                value = (String) object;
                if (value.indexOf("/") < 0 || value.startsWith("/") || value.endsWith("/")) {
                    throw new ExcelParseException("divideCellFormula value format is error");
                }
                String[] split = value.split("\\/");
                //拼接格式：B1/E1
                value = Arrays.stream(split).collect(Collectors.joining(row + "/", "", row + ""));
            }
        }
        return "IFERROR(" + value + ",\"\")";
    }
}
