package com.iris.excelfile.utils;

import com.iris.excelfile.constant.ExcelTypeEnum;
import com.iris.excelfile.exception.ExcelParseException;
import com.iris.excelfile.metadata.ExcelSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;


public class WorkBookUtil {


    public static Workbook createWorkBook(InputStream templateInputStream, ExcelTypeEnum excelType) throws IOException {
        Workbook workbook;
        if (ExcelTypeEnum.XLS.equals(excelType)) {
            workbook = (templateInputStream == null) ? new HSSFWorkbook() : new HSSFWorkbook(
                    new POIFSFileSystem(templateInputStream));
        } else {
//            workbook = (templateInputStream == null) ? new SXSSFWorkbook(500) : new SXSSFWorkbook(
//                    new XSSFWorkbook(templateInputStream));
            workbook = (templateInputStream == null) ? new SXSSFWorkbook(500) :
                    new XSSFWorkbook(templateInputStream);
        }
        return workbook;
    }

    public static Workbook createWorkBookWithStream(InputStream inputStream, ExcelTypeEnum excelType) throws IOException {
        if (inputStream == null) {
            throw new IOException("inputStream is null");
        }
        Workbook workbook;
        if (ExcelTypeEnum.XLS.equals(excelType)) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    public static Sheet createOrGetSheet(Workbook workbook, int sheetAt) {
        Sheet sheet = null;
        try {
            sheet = workbook.getSheetAt(sheetAt);
            if (sheet == null) {
                sheet = workbook.createSheet();
            }
        } catch (Exception e) {
            throw new ExcelParseException("create or get excel sheet fail");
        }
        return sheet;
    }

    public static Row createOrGetRow(Sheet sheet, int rowNum) {
//		synchronized (Object.class) {
        Row row = null;
        try {
            row = sheet.getRow(rowNum);
            if (row == null) {
                row = sheet.createRow(rowNum);
            }
        } catch (Exception e) {
            row = sheet.createRow(rowNum);
        }
        return row;
//		}
    }

    public static Row createRow(Sheet sheet, int rowNum) {
        return sheet.createRow(rowNum);
    }

    public static Cell createCell(Row row, int colNum, CellStyle cellStyle, String cellValue) {
        return createCell(row, colNum, cellStyle, cellValue, false, false, false);
    }

    public static Sheet createSheet(Workbook workbook, ExcelSheet excelSheet) {
        return workbook.createSheet(excelSheet.getSheetName() != null ? excelSheet.getSheetName() : excelSheet.getSheetNo() + "");
    }

    public static Cell createCell(Sheet sheet, Row row, int colNum, CellStyle cellStyle, Object cellValue, Boolean isNum, Boolean isFormula, boolean keepTpStyle, boolean isSeqNo) {
        Cell cell = null;
        if (!keepTpStyle || isSeqNo) {
            cell = row.createCell(colNum);
            cell.setCellStyle(cellStyle);
        } else if (sheet != null) {
            cell = sheet.getRow(row.getRowNum()).getCell(colNum);
        }
        if (null != cellValue && cell != null) {
            if (isFormula) {
                cell.setCellFormula(cellValue.toString());
            } else {
                if (isNum && !isSeqNo) {
                    cell.setCellValue(Double.parseDouble(cellValue.toString()));
                } else {
                    cell.setCellValue(cellValue.toString());
                }
            }
        }
        return cell;
    }

    public static Cell createCell(Row row, int colNum, CellStyle cellStyle, Object cellValue, Boolean isNum, Boolean isFormula, boolean isSeqNo) {
        return createCell(null, row, colNum, cellStyle, cellValue, isNum, isFormula, false, isSeqNo);
    }


}
