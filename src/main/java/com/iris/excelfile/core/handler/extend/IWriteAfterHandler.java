package com.iris.excelfile.core.handler.extend;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public interface IWriteAfterHandler {

    /**
     * 执行顺序1
     *
     * @param rowNum
     * @param row
     */
    void row(int rowNum, Row row);

    /**
     * 执执行顺序2
     *
     * @param cellNum
     * @param cell
     */
    void cell(int cellNum, Cell cell);
}
