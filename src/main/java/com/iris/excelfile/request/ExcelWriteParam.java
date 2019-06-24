package com.iris.excelfile.request;


import com.iris.excelfile.metadata.ExcelSheet;
import com.iris.excelfile.utils.FileUtil;

import java.io.Serializable;
import java.util.List;

/**
 * 导出文档请求参数
 *
 * @author liu_wp
 * @date Created in 2019/2/28 18:01
 * @see
 */
public class ExcelWriteParam implements Serializable {
    /**
     *
     */
    private static final long serialVersionUID = 1L;
    /**
     * excel 模板文件
     */
    private String excelTemplateFile;
    /**
     * excel 导出路径
     */
    private String excelOutFilePath;
    /**
     * excel 文件名称
     */
    private String excelFileName;
    /**
     * excel 导出全路径（路径+文件名称）
     */
    private String excelOutFileFullPath;
    /**
     * excel 加密密码 设置有值会加密文档
     */
    private String encryptPwd;
    /**
     * excel sheets
     */
    private List<ExcelSheet> excelSheets;

    public List<ExcelSheet> getExcelSheets() {
        return excelSheets;
    }

    public void setExcelSheets(List<ExcelSheet> excelSheets) {
        this.excelSheets = excelSheets;
    }

    public String getExcelTemplateFile() {
        return excelTemplateFile;
    }

    public void setExcelTemplateFile(String excelTemplateFile) {
        this.excelTemplateFile = excelTemplateFile;
    }

    public String getExcelOutFilePath() {
        return excelOutFilePath;
    }

    public void setExcelOutFilePath(String excelOutFilePath) {
        this.excelOutFilePath = excelOutFilePath;
    }

    public String getExcelFileName() {
        return excelFileName;
    }

    public void setExcelFileName(String excelFileName) {
        this.excelFileName = excelFileName;
    }

    public String getExcelOutFileFullPath() {
        if (this.excelOutFileFullPath != null && this.excelOutFileFullPath.length() > 0) {
            return this.excelOutFileFullPath;
        }
        return FileUtil.fileReplacePath(this.excelOutFilePath + "/" + this.excelFileName);
    }

    public void setExcelOutFileFullPath(String excelOutFileFullPath) {
        this.excelOutFileFullPath = excelOutFileFullPath;
    }

    public String getEncryptPwd() {
        return encryptPwd;
    }

    public void setEncryptPwd(String encryptPwd) {
        this.encryptPwd = encryptPwd;
    }
}
