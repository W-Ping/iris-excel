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
 * @author thinkpad
 * @date 2019/6/26 17:07
 * @see
 */
public class CrossTeamWriteBeforeHandlerImpl implements IWriteBeforeHandler {
    @Override
    public Map<Integer, CellStyle> customSetCellStyle(Workbook workbook, short dataFormat) {
        Map<Integer, CellStyle> map = new HashMap<>();
        CellStyle newCellStyle = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.WHITE, false);
        map.put(51, newCellStyle);
        return map;
    }

    @Override
    public boolean clearCellValue(Row row, int cellIndex) {
        return false;
    }
}
