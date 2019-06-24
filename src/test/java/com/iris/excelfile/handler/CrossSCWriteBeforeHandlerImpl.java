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
public class CrossSCWriteBeforeHandlerImpl implements IWriteBeforeHandler {
    /**
     * @param workbook
     * @return
     */
    @Override
    public Map<Integer, CellStyle> customSetCellStyle(Workbook workbook, short dataFormat) {
        Map<Integer, CellStyle> map = new HashMap<>();
        CellStyle newCellStyle = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.WHITE, true);
        map.put(0, newCellStyle);
        map.put(1, newCellStyle);
        map.put(3, newCellStyle);
        map.put(7, newCellStyle);
        map.put(11, newCellStyle);
        map.put(15, newCellStyle);
        map.put(17, newCellStyle);
        map.put(19, newCellStyle);
        map.put(21, newCellStyle);
        map.put(23, newCellStyle);
        map.put(27, newCellStyle);
        map.put(31, newCellStyle);
        map.put(47, newCellStyle);
        map.put(49, newCellStyle);
        map.put(51, newCellStyle);
        map.put(53, newCellStyle);
        map.put(55, newCellStyle);
        CellStyle newCellStyle1 = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.DARK_BLUE, true);
        map.put(5, newCellStyle1);
        map.put(6, newCellStyle1);
        map.put(9, newCellStyle1);
        map.put(10, newCellStyle1);
        map.put(13, newCellStyle1);
        map.put(14, newCellStyle1);
        return map;
    }

    /**
     * 清空列值
     *
     * @param cellIndex
     */
    @Override
    public boolean clearCellValue(Row row, int cellIndex) {
        return false;
    }
}
