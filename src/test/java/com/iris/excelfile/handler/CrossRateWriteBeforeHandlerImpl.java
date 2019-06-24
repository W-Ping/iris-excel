package com.iris.excelfile.handler;

import com.iris.excelfile.core.handler.extend.IWriteBeforeHandler;
import com.iris.excelfile.utils.StyleUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 * @author liu_wp
 * @date Created in 2019/6/20 11:58
 * @see
 */
public class CrossRateWriteBeforeHandlerImpl implements IWriteBeforeHandler {
    @Override
    public Map<Integer, CellStyle> customSetCellStyle(Workbook workbook, short dataFormat) {
        Map<Integer, CellStyle> map = new HashMap<>();
        CellStyle newCellStyle1 = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.WHITE, false);
        map.put(0, newCellStyle1);
        return map;
    }

    @Override
    public boolean clearCellValue(Row row, int cellIndex) {
        return false;
    }
}
