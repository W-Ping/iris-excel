package com.iris.excelfile.model;

import com.iris.excelfile.annotation.ExcelWriteProperty;
import com.iris.excelfile.metadata.BaseRowModel;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * @author liu_wp
 * @date Created in 2019/6/18 18:05
 * @see
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class CrossReportRateModel extends BaseRowModel {

    private String empType;
    private String dateVersion;
    @ExcelWriteProperty(index = 0)
    private String empCode;
    @ExcelWriteProperty(index = 1, mergeCellIndex = 2, divideCellFormula = "D/B")
    private String customerFlowAllTargetValue;
    @ExcelWriteProperty(index = 2)
    private String customerFlowAllActualValue;
    @ExcelWriteProperty(index = 3, mergeCellIndex = 4)
//    @ExcelWriteProperty(index = 3, isSeqNo = true, mergeCellIndex = 4, divideCellFormula = "E/C")
    private String customerFlowTargetValue;
    @ExcelWriteProperty(index = 4)
    private String customerFlowActualValue;
    @ExcelWriteProperty(index = 5, mergeCellIndex = 6, divideCellFormula = "H/F")
    private String incomingCallAllTargetValue;
    @ExcelWriteProperty(index = 6)
    private String incomingCallAllActualValue;
    @ExcelWriteProperty(index = 7, mergeCellIndex = 8, divideCellFormula = "I/G")
    private String incomingCallTargetValue;
    @ExcelWriteProperty(index = 8)
    private String incomingCallActualValue;
    @ExcelWriteProperty(index = 9, mergeCellIndex = 10, divideCellFormula = "L/J")
    private String networkAllTargetValue;
    @ExcelWriteProperty(index = 10)
    private String networkAllActualValue;
    @ExcelWriteProperty(index = 11, mergeCellIndex = 12, divideCellFormula = "M/K")
    private String networkTargetValue;
    @ExcelWriteProperty(index = 12)
    private String networkActualValue;
    @ExcelWriteProperty(index = 13, mergeCellIndex = 14, divideCellFormula = "P/N")
    private String initiativeCollectAllTargetValue;
    @ExcelWriteProperty(index = 14)
    private String initiativeCollectAllActualValue;
    @ExcelWriteProperty(index = 15, mergeCellIndex = 16, divideCellFormula = "Q/O")
    private String initiativeCollectTargetValue;
    @ExcelWriteProperty(index = 16)
    private String initiativeCollectActualValue;
    @ExcelWriteProperty(index = 17, mergeCellIndex = 18, divideCellFormula = "T/R")
    private String recommendAllTargetValue;
    @ExcelWriteProperty(index = 18)
    private String recommendAllActualValue;
    @ExcelWriteProperty(index = 19, mergeCellIndex = 20, divideCellFormula = "U/S")
    private String recommendTargetValue;
    @ExcelWriteProperty(index = 20)
    private String recommendActualValue;

    @ExcelWriteProperty(index = 21, mergeCellIndex = 22, divideCellFormula = "X/V")
    private String buyAgainAllTargetValue;
    @ExcelWriteProperty(index = 22)
    private String buyAgainAllActualValue;
    @ExcelWriteProperty(index = 23, mergeCellIndex = 24, divideCellFormula = "Y/W")
    private String buyAgainTargetValue;
    @ExcelWriteProperty(index = 24)
    private String buyAgainActualValue;
    @ExcelWriteProperty(index = 25, mergeCellIndex = 26, divideCellFormula = "AB/Z")
    private String activeAllTargetValue;
    @ExcelWriteProperty(index = 26)
    private String activeAllActualValue;
    @ExcelWriteProperty(index = 27, mergeCellIndex = 28, divideCellFormula = "AC/AA")
    private String activeTargetValue;
    @ExcelWriteProperty(index = 28)
    private String activeActualValue;
    @ExcelWriteProperty(index = 29, mergeCellIndex = 30, divideCellFormula = "AF/AD")
    private String dormantAllTargetValue;
    @ExcelWriteProperty(index = 30)
    private String dormantAllActualValue;
    @ExcelWriteProperty(index = 31, mergeCellIndex = 32, divideCellFormula = "AG/AE")
    private String dormantTargetValue;
    @ExcelWriteProperty(index = 32)
    private String dormantActualValue;
    @ExcelWriteProperty(index = 34, divideCellFormula = "AK/AI")
    private String customerFlowTotalTargetValue;
    @ExcelWriteProperty(index = 35, divideCellFormula = "AL/AJ")
    private String customerFlowTotalActualValue;
    @ExcelWriteProperty(index = 36, divideCellFormula = "AM/AK")
    private String intoStoreSaleTotalTargetValue;
    @ExcelWriteProperty(index = 37, divideCellFormula = "AN/AL")
    private String intoStoreSaleTotalActualValue;
    @ExcelWriteProperty(index = 38, divideCellFormula = "AO/AM")
    private String quotedPriceCountTargetValue;
    @ExcelWriteProperty(index = 39, divideCellFormula = "AP/AN")
    private String quotedPriceCountActualValue;
    @ExcelWriteProperty(index = 40, divideCellFormula = "AQ/AO")
    private String orderCountTargetValue;
    @ExcelWriteProperty(index = 41, divideCellFormula = "AR/AP")
    private String orderCountActualValue;
    @ExcelWriteProperty(index = 42, divideCellFormula = "AO/AK")
    private String invoicedCountTargetValue;
    @ExcelWriteProperty(index = 43, divideCellFormula = "AP/AL")
    private String invoicedCountActualValue;
    @ExcelWriteProperty(index = 44, divideCellFormula = "AS/AK")
    private String firstTestDriveCountTargetValue;
    @ExcelWriteProperty(index = 45, divideCellFormula = "AT/AL")
    private String firstTestDriveCountActualValue;
//    @ExcelWriteProperty(index = 46, dateFormat = "YYYY年MM月dd日")
//    private Date date;
//    @ExcelWriteProperty(index = 47, dateFormat = "YYYY-MM-dd HH:mm:ss")
//    private String date2;
//    @ExcelWriteProperty(index = 48)
//    private String phone;
//    @ExcelWriteProperty(index = 49)
//    private Integer phone2;
}
