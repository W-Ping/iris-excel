package com.iris.excelfile.response;

import com.iris.excelfile.constant.FileConstant;
import org.apache.poi.ss.formula.functions.T;

/**
 * @author liu_wp
 * @date 2019/8/30 9:29
 * @see
 */
public class ExcelReadResponse extends BaseResponse {
    private Object data;

    public ExcelReadResponse() {
        super();
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }
}
