package com.iris.excelfile.data;

import com.iris.excelfile.core.handler.extend.DictionaryRefHandler;

/**
 * @author thinkpad
 * @date 2019/6/28 17:35
 * @see
 */
public class TemplateDic extends DictionaryRefHandler {
    @Override
    protected String getDicValue(Integer cellIndex, Integer code) {
        if (cellIndex == 2) {
            return TemplateEum.getTypeName(code);
        }
        return null;
    }
}
