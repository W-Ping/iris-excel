package com.iris.excelfile.model;

import com.iris.excelfile.annotation.ExcelWriteProperty;
import com.iris.excelfile.metadata.BaseRowModel;
import lombok.Data;
import lombok.EqualsAndHashCode;

@Data
@EqualsAndHashCode(callSuper = false)
public class CrossReportRateGroupNameModel extends BaseRowModel {
    private String empType;
    private String dateVersion;
    @ExcelWriteProperty(index = 0)
    private String empCode;
    @ExcelWriteProperty(index = 1)
    private Integer customerFlowAllTargetValue;
    @ExcelWriteProperty(index = 3)
    private Integer customerFlowTargetValue;
    @ExcelWriteProperty(index = 2)
    private Integer customerFlowAllActualValue;
    @ExcelWriteProperty(index = 4)
    private Integer customerFlowActualValue;
    @ExcelWriteProperty(index = 5)
    private Integer incomingCallAllTargetValue;
    @ExcelWriteProperty(index = 6)
    private Integer incomingCallAllActualValue;
    @ExcelWriteProperty(index = 7)
    private Integer incomingCallTargetValue;
    @ExcelWriteProperty(index = 8)
    private Integer incomingCallActualValue;
    @ExcelWriteProperty(index = 9)
    private Integer networkAllTargetValue;
    @ExcelWriteProperty(index = 10)
    private Integer networkAllActualValue;
    @ExcelWriteProperty(index = 11)
    private Integer networkTargetValue;
    @ExcelWriteProperty(index = 12)
    private Integer networkActualValue;
    @ExcelWriteProperty(index = 13)
    private Integer initiativeCollectAllTargetValue;
    @ExcelWriteProperty(index = 14)
    private Integer initiativeCollectAllActualValue;
    @ExcelWriteProperty(index = 15)
    private Integer initiativeCollectTargetValue;
    @ExcelWriteProperty(index = 16)
    private Integer initiativeCollectActualValue;
    @ExcelWriteProperty(index = 17)
    private Integer recommendAllTargetValue;
    @ExcelWriteProperty(index = 18)
    private Integer recommendAllActualValue;
    @ExcelWriteProperty(index = 19)
    private Integer recommendTargetValue;
    @ExcelWriteProperty(index = 20)
    private Integer recommendActualValue;
    @ExcelWriteProperty(index = 21)
    private Integer buyAgainAllTargetValue;
    @ExcelWriteProperty(index = 22)
    private Integer buyAgainAllActualValue;
    @ExcelWriteProperty(index = 23)
    private Integer buyAgainTargetValue;
    @ExcelWriteProperty(index = 24)
    private Integer buyAgainActualValue;
    @ExcelWriteProperty(index = 25)
    private Integer activeAllTargetValue;
    @ExcelWriteProperty(index = 26)
    private Integer activeAllActualValue;
    @ExcelWriteProperty(index = 27)
    private Integer activeTargetValue;
    @ExcelWriteProperty(index = 28)
    private Integer activeActualValue;
    @ExcelWriteProperty(index = 29)
    private Integer dormantAllTargetValue;
    @ExcelWriteProperty(index = 30)
    private Integer dormantAllActualValue;
    @ExcelWriteProperty(index = 31)
    private Integer dormantTargetValue;
    @ExcelWriteProperty(index = 32)
    private Integer dormantActualValue;
    @ExcelWriteProperty(index = 34)
    private Integer customerFlowTotalTargetValue;
    @ExcelWriteProperty(index = 35)
    private Integer customerFlowTotalActualValue;
    @ExcelWriteProperty(index = 36)
    private Integer intoStoreSaleTotalTargetValue;
    @ExcelWriteProperty(index = 37)
    private Integer intoStoreSaleTotalActualValue;
    @ExcelWriteProperty(index = 38)
    private Integer quotedPriceCountTargetValue;
    @ExcelWriteProperty(index = 39)
    private Integer quotedPriceCountActualValue;
    @ExcelWriteProperty(index = 40)
    private Integer orderCountTargetValue;
    @ExcelWriteProperty(index = 41)
    private Integer orderCountActualValue;
    @ExcelWriteProperty(index = 42)
    private Integer invoicedCountTargetValue;
    @ExcelWriteProperty(index = 43)
    private Integer invoicedCountActualValue;
    @ExcelWriteProperty(index = 44)
    private Integer firstTestDriveCountTargetValue;
    @ExcelWriteProperty(index = 45)
    private Integer firstTestDriveCountActualValue;
}
