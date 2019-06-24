package com.iris.excelfile.handler;

import com.iris.excelfile.core.handler.extend.IWriteAfterHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author liu_wp
 * @date Created in 2019/3/29 15:46
 * @see
 */
public class WriteAfterHandlerImpl implements IWriteAfterHandler {

    @Override
    public void row(int rowNum, Row row) {
    }

    @Override
    public void cell(int cellNum, Cell cell) {
        if (cellNum == 5 || cellNum == 6 || cellNum == 9 || cellNum == 10 || cellNum == 13 || cellNum == 14) {
//            CellStyle cellStyle = cell.getCellStyle();
//            cellStyle.setFillBackgroundColor(IndexedColors.BLUE.index);
        }
    }
}
