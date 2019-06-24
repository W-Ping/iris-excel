package com.iris.excelfile.response;


import com.iris.excelfile.constant.FileConstant;

import java.io.Serializable;

/**
 * @author liu_wp
 * @date Created in 2019/3/12 11:17
 * @see
 */
public class BaseResponse implements Serializable {
    /**
     *
     */
    private static final long serialVersionUID = 1L;
    /**
     *
     */
    private String code;
    /**
     *
     */
    private String message;

    public BaseResponse(String code, String message) {
        this.code = code;
        this.message = message;
    }

    public BaseResponse() {
        this.code = FileConstant.FAIL_CODE;
        this.message = "error";
    }


    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

}
