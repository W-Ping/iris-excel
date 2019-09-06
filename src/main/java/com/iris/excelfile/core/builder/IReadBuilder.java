package com.iris.excelfile.core.builder;

import com.iris.excelfile.metadata.ExcelReadColumnProperty;
import org.apache.poi.ss.formula.functions.T;

import java.util.List;
import java.util.Map;

/**
 * @author liu_wp
 * @date Created in 2019/8/30 9:21
 * @see
 */
public interface IReadBuilder {
    /**
     * @param classes
     * @return
     */
    void initDataClass(List<Class> classes);

    /**
     *
     */
    Map<Integer, Object> addExcelToObject();

    /**
     * @return
     */
    Object addOneSheetExcelToObject();
}
