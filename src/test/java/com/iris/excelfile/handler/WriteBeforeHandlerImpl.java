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
 * @date Created in 2019/3/29 15:55
 * @see
 */
public class WriteBeforeHandlerImpl implements IWriteBeforeHandler {
    @Override
    public Map<Integer, CellStyle> customSetCellStyle(Workbook workbook, short dataFormat) {
        Map<Integer, CellStyle> map = new HashMap<>();
        CellStyle newCellStyle = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.DARK_YELLOW, true);
        map.put(0, newCellStyle);
        map.put(2, newCellStyle);
        map.put(4, newCellStyle);
        return map;
    }

    /**
     * 清空列值
     *
     * @param row
     * @param cellIndex
     */
    @Override
    public boolean clearCellValue(Row row, int cellIndex) {
        return false;
    }
}
