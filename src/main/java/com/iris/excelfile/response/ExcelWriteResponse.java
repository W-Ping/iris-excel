package com.iris.excelfile.response;


import com.iris.excelfile.constant.FileConstant;

/**
 * @author liu_wp
 * @date Created in 2019/3/12 11:16
 * @see
 */
public class ExcelWriteResponse extends BaseResponse {

    private String excelOutFilePath;

    public ExcelWriteResponse(String code, String message, String excelOutFilePath) {
        super(code, message);
        this.excelOutFilePath = excelOutFilePath;
    }

    public ExcelWriteResponse(String code, String excelOutFilePath) {
        this(code, null, excelOutFilePath);
    }

    public static ExcelWriteResponse fail(String error) {
        ExcelWriteResponse excelWriteResponse = new ExcelWriteResponse();
        excelWriteResponse.setCode(FileConstant.FAIL_CODE);
        excelWriteResponse.setMessage(error);
        return excelWriteResponse;
    }

    public static ExcelWriteResponse success(String excelOutFilePath) {
        ExcelWriteResponse excelWriteResponse = new ExcelWriteResponse();
        excelWriteResponse.setCode(FileConstant.SUCCESS_CODE);
        excelWriteResponse.setExcelOutFilePath(excelOutFilePath);
        return excelWriteResponse;
    }

    public ExcelWriteResponse() {
        super();
    }

    public String getExcelOutFilePath() {
        return excelOutFilePath;
    }

    public void setExcelOutFilePath(String excelOutFilePath) {
        this.excelOutFilePath = excelOutFilePath;
    }
}
