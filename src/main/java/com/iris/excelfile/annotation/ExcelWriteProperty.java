package com.iris.excelfile.annotation;

import java.lang.annotation.*;

/**
 * 写入Excel注解
 *
 * @author liu_wp
 * @date Created in 2019/3/5 10:46
 * @see
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelWriteProperty {

    String[] head() default {};

    int index() default 99999;

    int mergeCellIndex() default -1;

    boolean isSeqNo() default false;

    /**
     * @return
     */
    int mergeRowCount() default 0;

    /**
     * @return
     */
    String[] sumCellFormula() default {};

    /**
     * @return
     */
    String divideCellFormula() default "";

    /**
     * @return
     */
    boolean ignore() default false;

    /**
     * @return
     */
    String dateFormat() default "";

}
