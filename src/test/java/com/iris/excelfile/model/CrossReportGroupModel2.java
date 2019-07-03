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
    @ExcelWriteProperty(index = 0, ignoreCell = true)
    private String empCode;
    @ExcelWriteProperty(index = 47, ignoreCell = true)
    private Integer financialCountTargetValue;
    @ExcelWriteProperty(index = 48, ignoreCell = true)
    private Integer financialCountActualValue;
    @ExcelWriteProperty(index = 49, ignoreCell = true)
    private Integer insuranceCountTargetValue;
    @ExcelWriteProperty(index = 50, ignoreCell = true)
    private Integer insuranceCountActualValue;
    @ExcelWriteProperty(index = 51, ignoreCell = true)
    private BigDecimal skuGoodsAmountTargetValue;
    @ExcelWriteProperty(index = 52, ignoreCell = true)
    private BigDecimal skuGoodsAmountActualValue;
    @ExcelWriteProperty(index = 53, ignoreCell = true)
    private Integer extendWarrantyCountTargetValue;
    @ExcelWriteProperty(index = 54, ignoreCell = true)
    private Integer extendWarrantyCountActualValue;
    @ExcelWriteProperty(index = 55, ignoreCell = true)
    private Integer otherCountTargetValue;
    @ExcelWriteProperty(index = 56, ignoreCell = true)
    private Integer otherCountActualValue;
}
