package com.iris.excelfile.exception;

/**
 * @author liu_wp
 * @date Created in 2019/6/25 12:52
 * @see
 */
public class ExcelException extends Exception {
    public ExcelException() {
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, String error) {
        super(message + "，【" + error + "】");
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

}
