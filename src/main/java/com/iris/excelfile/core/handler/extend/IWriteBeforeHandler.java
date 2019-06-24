package com.iris.excelfile.core.handler.extend;

import com.iris.excelfile.metadata.ExcelTable;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * @author liu_wp
 * @date Created in 2019/3/29 14:57
 * @see
 */
public interface IWriteBeforeHandler {
    /**
     * 自定义设置列样式
     *
     * @param workbook
     * @param dataFormat
     * @return
     */
    Map<Integer, CellStyle> customSetCellStyle(Workbook workbook, short dataFormat);
    
    /**
     * 清空列值
     *
     * @param row
     * @param cellIndex
     * @return
     */
    boolean clearCellValue(Row row, int cellIndex);
}
