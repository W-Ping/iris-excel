package com.iris.excelfile.request;


import org.apache.poi.ss.formula.functions.T;

import java.io.Serializable;
import java.util.List;

/**
 * @author liu_wp
 * @date 2019/8/30 9:32
 * @see
 */
public class ExcelReadParam implements Serializable {
    /**
     *
     */
    private static final long serialVersionUID = 1L;
    private List<Class> dataClass;
    private String excelPath;


    public String getExcelPath() {
        return excelPath;
    }

    public List<Class> getDataClass() {
        return dataClass;
    }

    public void setDataClass(List<Class> dataClass) {
        this.dataClass = dataClass;
    }

    public void setExcelPath(String excelPath) {
        this.excelPath = excelPath;
    }
}
