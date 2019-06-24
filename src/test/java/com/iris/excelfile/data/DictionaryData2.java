package com.iris.excelfile.data;

import com.iris.excelfile.core.handler.DictionaryRefHandler;

public class DictionaryData2 extends DictionaryRefHandler {
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
