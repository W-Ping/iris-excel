package com.iris.excelfile.data;

import com.iris.excelfile.core.handler.DictionaryRefHandler;

public class DictionaryData extends DictionaryRefHandler {
    /**
     * @param i
     * @param code
     * @return
     */
    @Override
    protected String getDicValue(Integer cellIndex, Integer code) {
        if (cellIndex == 4 || cellIndex == 6) {
            return Dictionary.getDictionaryValue(code);
        }
        return null;
    }
}
