package com.iris.excelfile.annotation;

import java.lang.annotation.*;

/**
 * @author liu_wp
 * @date Created in 2019/8/30 9:13
 * @see
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelReadProperty {
    /**
     * @return
     */
    int cellIndex();

    /**
     * @return
     */
    String dateFormat() default "";

    /**
     * @return
     */
    String numberFormat() default "";
}
