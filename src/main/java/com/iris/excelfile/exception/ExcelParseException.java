package com.iris.excelfile.exception;

import org.apache.commons.lang3.StringUtils;

/**
 * @author liu_wp
 * @date Created in 2019/3/1 10:07
 * @see
 */
public class ExcelParseException extends RuntimeException {

    /**
     *
     */
    public ExcelParseException() {
        super();
    }

    /**
     * @param message
     */
    public ExcelParseException(String message) {
        super(message);
    }

    /**
     * @param message
     * @param cause
     */
    public ExcelParseException(String message, Throwable cause) {
        super(message + "，【" + (StringUtils.isNotBlank(cause.getMessage()) ? cause.getMessage() : cause.toString()) + "】", cause);
    }

    /**
     * @param cause
     */
    public ExcelParseException(Throwable cause) {
        super(cause);
    }
}
