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
 * @date Created in 2019/6/17 17:43
 * @see
 */
public class CrossDccWriteBeforeHandlerImpl implements IWriteBeforeHandler {
    /**
     * 设置table 列样式
     *
     * @param workbook
     * @return
     */
    @Override
    public Map<Integer, CellStyle> customSetCellStyle(Workbook workbook, short dataFormat) {
        Map<Integer, CellStyle> map = new HashMap<>();
        CellStyle newCellStyle1 = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.WHITE, false);
        newCellStyle1.setDataFormat(dataFormat);
        map.put(0, newCellStyle1);
        map.put(5, newCellStyle1);
        map.put(7, newCellStyle1);
        map.put(9, newCellStyle1);
        map.put(11, newCellStyle1);
        map.put(15, newCellStyle1);
        CellStyle newCellStyle = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.DARK_BLUE, true);
        map.put(47, newCellStyle);
        map.put(48, newCellStyle);
        map.put(49, newCellStyle);
        map.put(50, newCellStyle);
        map.put(51, newCellStyle);
        map.put(52, newCellStyle);
        map.put(53, newCellStyle);
        map.put(54, newCellStyle);
        map.put(55, newCellStyle);
        map.put(56, newCellStyle);
        return map;
    }

    /**
     * 清空列值
     *
     * @param cellIndex
     */
    @Override
    public boolean clearCellValue(Row row, int cellIndex) {
//        int rowNum = row.getRowNum();
//        if (cellIndex == 1) {
//            return true;
//        }
        return false;
    }
}
