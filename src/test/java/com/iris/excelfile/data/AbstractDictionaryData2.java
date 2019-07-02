package com.iris.excelfile.data;

import com.iris.excelfile.core.handler.extend.AbstractDictionaryRefHandler;

public class AbstractDictionaryData2 extends AbstractDictionaryRefHandler {
    /**
     * @param i
     * @param code
     * @return
     */
    @Override
    protected String getDicValue(Integer cellIndex, Integer code) {
        if (cellIndex == 1 || cellIndex == 5) {
            return Dictionary2.getDictionaryValue(code);
        }
        return null;
    }
}
