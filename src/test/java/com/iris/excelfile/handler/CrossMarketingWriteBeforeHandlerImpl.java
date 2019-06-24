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
 * @date Created in 2019/6/18 17:31
 * @see
 */
public class CrossMarketingWriteBeforeHandlerImpl implements IWriteBeforeHandler {
    /**
     * 设置table 列样式
     *
     * @param workbook
     * @return
     */
    @Override
    public Map<Integer, CellStyle> customSetCellStyle(Workbook workbook, short dataFormat) {
        Map<Integer, CellStyle> map = new HashMap<>();
        CellStyle newCellStyle = StyleUtil.buildBaseIsNotBoldCellStyle(workbook, null, IndexedColors.WHITE, false);
        map.put(1, newCellStyle);
        map.put(13, newCellStyle);
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
