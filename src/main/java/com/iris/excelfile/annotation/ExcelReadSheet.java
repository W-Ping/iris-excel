package com.iris.excelfile.annotation;

import java.lang.annotation.*;

/**
 * @author liu_wp
 * @date Created in 2019/8/30 10:45
 * @see
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelReadSheet {
    /**
     * @return
     */
    int SheetNo() default 0;

    /**
     * 开始行数
     *
     * @return
     */
    int startRowIndex() default 1;

    /**
     * 开始列数
     *
     * @return
     */
    int startCellIndex() default 0;
}
