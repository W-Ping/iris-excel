package com.iris.excelfile.model;

import com.iris.excelfile.annotation.ExcelWriteProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.math.BigDecimal;

/**
 * @author liu_wp
 * @date Created in 2019/6/18 16:15
 * @see
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class CrossReportGroupModel2 extends CrossReportModel {
    @ExcelWriteProperty(index = 0, ignore = true)
    private String empCode;
    @ExcelWriteProperty(index = 47, keepTpStyle = true)
    private Integer financialCountTargetValue;
    @ExcelWriteProperty(index = 48, keepTpStyle = true)
    private Integer financialCountActualValue;
    @ExcelWriteProperty(index = 49, keepTpStyle = true)
    private Integer insuranceCountTargetValue;
    @ExcelWriteProperty(index = 50, keepTpStyle = true)
    private Integer insuranceCountActualValue;
    @ExcelWriteProperty(index = 51, keepTpStyle = true)
    private BigDecimal skuGoodsAmountTargetValue;
    @ExcelWriteProperty(index = 52, keepTpStyle = true)
    private BigDecimal skuGoodsAmountActualValue;
    @ExcelWriteProperty(index = 53, keepTpStyle = true)
    private Integer extendWarrantyCountTargetValue;
    @ExcelWriteProperty(index = 54, keepTpStyle = true)
    private Integer extendWarrantyCountActualValue;
    @ExcelWriteProperty(index = 55, keepTpStyle = true)
    private Integer otherCountTargetValue;
    @ExcelWriteProperty(index = 56, keepTpStyle = true)
    private Integer otherCountActualValue;
}
