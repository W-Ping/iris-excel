package com.iris.excelfile.exception;

/**
 * @author liu_wp
 * @date Created in 2019/3/1 10:07
 * @see
 */
public class ExcelParseException extends RuntimeException {

    public ExcelParseException() {
        super();
    }

    public ExcelParseException(String message) {
        super(message);
    }

    public ExcelParseException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelParseException(Throwable cause) {
        super(cause);
    }
}
