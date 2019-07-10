package com.iris.excelfile.request;


import com.iris.excelfile.metadata.ExcelSheet;
import com.iris.excelfile.utils.FileUtil;
import org.apache.commons.lang3.StringUtils;

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
     * excel 导出文档路径
     */
    private String excelOutFilePath;
    /**
     * excel 导出文件名称
     */
    private String excelFileName;
    /**
     * excel 导出文档全路径（导出文档路径+导出文件名称）
     */
    private String excelOutFileFullPath;
    /**
     * excel 加密文档密码
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
        if (StringUtils.isNotBlank(this.excelOutFilePath) && StringUtils.isNotBlank(this.excelFileName)) {
            return FileUtil.fileReplacePath(this.excelOutFilePath + "/" + this.excelFileName);
        }
        return StringUtils.EMPTY;
    }

    public void setExcelOutFileFullPath(String excelOutFileFullPath) {
        if (StringUtils.isNotBlank(excelOutFileFullPath)) {
            if (StringUtils.isBlank(this.excelOutFilePath)) {
                setExcelOutFilePath(FileUtil.getFileDir(excelOutFileFullPath));
            }
            if (StringUtils.isBlank(this.excelFileName)) {
                setExcelFileName(FileUtil.getFileName(excelOutFileFullPath));
            }
        }
        this.excelOutFileFullPath = excelOutFileFullPath;
    }

    public String getEncryptPwd() {
        return encryptPwd;
    }

    public void setEncryptPwd(String encryptPwd) {
        this.encryptPwd = encryptPwd;
    }
}
