package com.iris.excelfile.data;

import com.iris.excelfile.core.handler.extend.AbstractDictionaryRefHandler;

public class AbstractDictionaryData extends AbstractDictionaryRefHandler {
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
