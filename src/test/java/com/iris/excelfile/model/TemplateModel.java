package com.iris.excelfile.model;

import com.iris.excelfile.annotation.ExcelWriteProperty;
import com.iris.excelfile.metadata.BaseRowModel;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author thinkpad
 * @date 2019/6/28 17:20
 * @see
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class TemplateModel extends BaseRowModel {
    @ExcelWriteProperty(index = 0, isSeqNo = true)
    private int seq;
    @ExcelWriteProperty(index = 1)
    private String empName;
    @ExcelWriteProperty(index = 2)
    private Integer type;
    @ExcelWriteProperty(index = 3)
    private BigDecimal money;
    @ExcelWriteProperty(index = 4, dateFormat = "YYYY年MM HH时mm分ss秒")
    private Date date;
}
