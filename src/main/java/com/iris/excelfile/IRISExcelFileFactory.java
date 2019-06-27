package com.iris.excelfile;

import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.constant.FileConstant;
import com.iris.excelfile.core.handler.ExportWriteHandler;
import com.iris.excelfile.metadata.ExcelSheet;
import com.iris.excelfile.metadata.ExcelTable;
import com.iris.excelfile.request.ExcelWriteParam;
import com.iris.excelfile.response.ExcelWriteResponse;
import com.iris.excelfile.utils.ExcelV2007Util;
import com.iris.excelfile.utils.FileUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.InputStream;
import java.io.OutputStream;
import java.rmi.server.ExportException;
import java.util.List;

/**
 * @author liu_wp
 * @date Created in 2019/6/20 18:11
 * @see
 */
@Slf4j
public class IRISExcelFileFactory {

    public static void main(String[] args) {

    }


    /**
     * 导出-模板
     *
     * @param param
     * @return
     */
    public static ExcelWriteResponse exportV2007WithTemplate(ExcelWriteParam param) {
        return exportV2007WithTemplate(param, null, false);
    }

    /**
     * 导出-模板（多线程）
     *
     * @param param
     * @return
     */
    public static ExcelWriteResponse exportV2007QueueWithTemplate(ExcelWriteParam param) {
        return exportV2007WithTemplate(param, null, true);
    }

    /**
     * 导出-模板（多线程）
     *
     * @param param
     * @param inputStream
     * @return
     */
    public static ExcelWriteResponse exportV2007QueueWithTemplate(ExcelWriteParam param, InputStream inputStream) {
        return exportV2007WithTemplate(param, inputStream, true);
    }

    /**
     * 导出-模板（多线程）
     *
     * @param param
     * @param isQueue
     * @return
     */
    public static ExcelWriteResponse exportV2007WithTemplate(ExcelWriteParam param, InputStream inputStream) {
        return exportV2007WithTemplate(param, inputStream, false);
    }

    /**
     * 导出
     *
     * @param param
     * @return
     */
    public static ExcelWriteResponse exportV2007(ExcelWriteParam param) {
        return exportV2007(param, null, false);
    }

    /**
     * 导出
     *
     * @param param
     * @return
     */
    public static ExcelWriteResponse exportV2007Queue(ExcelWriteParam param) {
        return exportV2007(param, null, true);
    }

    /**
     * @param param
     * @param inputStream
     * @return
     */
    private static ExcelWriteResponse exportV2007Queue(ExcelWriteParam param, InputStream inputStream) {
        return exportV2007(param, inputStream, true);
    }

    /**
     * @param param
     * @param inputStream
     * @return
     */
    private static ExcelWriteResponse exportV2007(ExcelWriteParam param, InputStream inputStream) {
        return exportV2007(param, inputStream, false);
    }

    /**
     * @param param
     * @param isQueue
     * @return
     */
    private static ExcelWriteResponse exportV2007(ExcelWriteParam param, boolean isQueue) {
        return exportV2007(param, null, isQueue);
    }

    /**
     * @param param
     * @param outputStream
     * @return
     */
    private static ExcelWriteResponse exportV2007(ExcelWriteParam param, InputStream inputStream, boolean isQueue) {
        long startTime = System.currentTimeMillis();
        log.info("开始导出文档{}", startTime);
        ExcelWriteResponse excelWriteResponse = new ExcelWriteResponse();
        OutputStream outputStream = null;
        try {
            validateParam(param, ExcelTypeEnum.XLSX);
            if (inputStream == null && StringUtils.isNotBlank(param.getExcelTemplateFile())) {
                inputStream = FileUtil.synGetResourcesFileInputStream(param.getExcelTemplateFile());
            }
            outputStream = FileUtil.synGetResourcesFileOutputStream(param.getExcelOutFileFullPath());
            ExportWriteHandler exportHandler = new ExportWriteHandler(inputStream, outputStream, param.getExcelOutFileFullPath(), ExcelTypeEnum.XLSX);
            exportHandler.exportExcelV2007(param.getExcelSheets(), isQueue);
            //文档加密
            if (StringUtils.isNotBlank(param.getExcelOutFileFullPath()) && StringUtils.isNotBlank(param.getEncryptPwd())) {
                ExcelV2007Util.encrypt(param.getExcelOutFileFullPath(), param.getEncryptPwd());
            }
            excelWriteResponse = new ExcelWriteResponse(FileConstant.SUCCESS_CODE, param.getExcelOutFileFullPath());
        } catch (Exception e) {
            log.error("导出文件错误,{},{}", param.getExcelFileName(), e.getMessage());
            excelWriteResponse.setMessage("【" + param.getExcelFileName() + "】导出文件错误" + (e.getMessage() != null ? e.getMessage() : ""));
        } finally {
            FileUtil.close(null, outputStream);
        }
        long totalTime = (System.currentTimeMillis() - startTime) / 1000;
        log.info("文档导出结束{},耗时：{}秒", System.currentTimeMillis(), totalTime);
        return excelWriteResponse;
    }


    /**
     * @param param
     * @param inputStream
     * @return
     */
    private static ExcelWriteResponse exportV2007WithTemplate(ExcelWriteParam param, InputStream inputStream, boolean isQueue) {
        if (StringUtils.isBlank(param.getExcelTemplateFile())) {
            return ExcelWriteResponse.fail("excel export template is null");
        }
        if (inputStream == null) {
            FileUtil.checkFilePath(param.getExcelTemplateFile());
        }
        return exportV2007(param, inputStream, isQueue);
    }

    private static void validateParam(ExcelWriteParam excelWriteParam, ExcelTypeEnum excelTypeEnum) throws ExportException {
        if (excelWriteParam == null) {
            throw new ExportException("export params is null");
        }
        //如果没有传入导出文件信息 创建临时导出文件
        if (StringUtils.isBlank(excelWriteParam.getExcelOutFileFullPath())) {
            excelWriteParam.setExcelOutFileFullPath(FileUtil.createTempOutFile(FileUtil.createTempDirectory(), ExcelTypeEnum.XLSX));
        }
        String excelOutFileFullPath = excelWriteParam.getExcelOutFileFullPath();
        if (StringUtils.isBlank(excelOutFileFullPath)) {
            throw new ExportException("excel file is null");
        } else {
            if (excelOutFileFullPath.lastIndexOf(".") != -1) {
                if (ExcelTypeEnum.XLSX.equals(excelTypeEnum)
                        && !ExcelTypeEnum.XLSX.getValue().equals(excelOutFileFullPath.substring(excelOutFileFullPath.lastIndexOf(".")))) {
                    throw new ExportException("excel fileType is not xlsx");
                } else if (ExcelTypeEnum.XLS.equals(excelTypeEnum)
                        && !ExcelTypeEnum.XLS.getValue().equals(excelOutFileFullPath.substring(excelOutFileFullPath.lastIndexOf(".")))) {
                    throw new ExportException("excel fileType is not xls");
                }
            } else {
                throw new ExportException("excel file type is error");
            }
        }
        validateSheets(excelWriteParam.getExcelSheets());
    }


    private static void validateSheets(List<ExcelSheet> excelSheets) throws ExportException {
        if (CollectionUtils.isEmpty(excelSheets)) {
            throw new ExportException("excel sheet is null");
        }
        int sheetNo = -999;
        for (ExcelSheet sheet : excelSheets) {
            if (sheetNo == sheet.getSheetNo()) {
                throw new ExportException("sheetNo is repetition");
            }
            sheetNo = sheet.getSheetNo();
            validateTables(sheet.getExcelTables());
        }
    }

    private static void validateTables(List<ExcelTable> tables) throws ExportException {
        if (!CollectionUtils.isEmpty(tables)) {
            int tableNo = -999;
            for (ExcelTable table : tables) {
                if (tableNo == table.getTableNo()) {
                    throw new ExportException("tableNo " + tableNo + " is repetition");
                }
                tableNo = table.getTableNo();
                if (CollectionUtils.isEmpty(table.getHead()) && null == table.getHeadClass()) {
                    throw new ExportException("table head arguments is null");
                }
            }
        }
    }
}
