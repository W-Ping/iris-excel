package com.iris.excelfile.core.handler.extend;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public interface IWriteAfterHandler {

    /**
     * 执行顺序1
     *
     * @param rowNum 行索引
     * @param row    行
     */
    void row(int rowNum, Row row);

    /**
     * 执执行顺序2
     *
     * @param cellNum 列索引
     * @param cell    列
     */
    void cell(int cellNum, Cell cell);
}
