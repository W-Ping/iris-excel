package com.iris.excelfile.model;

import com.iris.excelfile.annotation.ExcelWriteProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * @author liu_wp
 * @date Created in 2019/6/18 16:15
 * @see
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class CrossReportGroupModel extends CrossReportModel {
    @ExcelWriteProperty(index = 0, ignoreCell = true)
    private String empCode;
}
