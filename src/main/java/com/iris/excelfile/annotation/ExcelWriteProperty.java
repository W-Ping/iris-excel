package com.iris.excelfile.annotation;

import java.lang.annotation.*;

/**
 * 写入Excel注解
 * 优先级别 ignoreCell>isSeqNo> sumCellFormula >divideCellFormula
 *
 * @author liu_wp
 * @date Created in 2019/3/5 10:46
 * @see
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelWriteProperty {

    /**
     * 表头
     *
     * @return
     */
    String[] head() default {};

    /**
     * 索引
     *
     * @return
     */
    int index() default 99999;

    /**
     * 合并列索引
     *
     * @return
     */
    int mergeCellIndex() default -1;

    /**
     * 合并的行数
     *
     * @return
     */
    int mergeRowCount() default 0;

    /**
     * 忽略当前列
     *
     * @return
     */
    boolean ignoreCell() default false;

    /**
     * 是否为序列号
     *
     * @return
     */
    boolean isSeqNo() default false;


    /**
     * SUM合计公式
     *
     * @return
     */
    String[] sumCellFormula() default {};

    /**
     * 除法 公式
     *
     * @return
     */
    String divideCellFormula() default "";


    /**
     * 日期格式
     *
     * @return
     */
    String dateFormat() default "";

}
