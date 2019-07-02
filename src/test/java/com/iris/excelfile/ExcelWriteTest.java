package com.iris.excelfile;

import com.iris.excelfile.constant.DataFormatEnum;
import com.iris.excelfile.constant.FileConstant;
import com.iris.excelfile.data.AbstractDictionaryData;
import com.iris.excelfile.data.AbstractDictionaryData2;
import com.iris.excelfile.data.TemplateDic;
import com.iris.excelfile.data.TestData;
import com.iris.excelfile.handler.*;
import com.iris.excelfile.metadata.*;
import com.iris.excelfile.model.*;
import com.iris.excelfile.request.ExcelWriteParam;
import com.iris.excelfile.response.ExcelWriteResponse;
import com.iris.excelfile.utils.ExcelV2007Util;
import com.iris.excelfile.utils.FileUtil;
import com.iris.excelfile.utils.JSONUtil;
import org.apache.poi.ss.usermodel.*;
import org.junit.Assert;
import org.junit.Test;
import org.springframework.util.CollectionUtils;

import java.io.InputStream;
import java.time.LocalDate;
import java.util.*;

public class ExcelWriteTest {
	/**
	 *
	 */
	@Test
	public void export2007WithTemplate() {
		String outFilePath = "D:\\导出EXCEL";
		String outFileName = ExcelV2007Util.generateUniqueFileNameByTime("跨部门月报") + ".xlsx";
		String outTemplate = "exportTemplateCrossDeptToMonth.xlsx";
		String excelOutFileFullPath = outFilePath + "\\" + outFileName;
		InputStream resourcesFileInputStream = FileUtil.getResourcesFileInputStream(outTemplate);
//        String outTemplate = "D:\\导出模板\\exportTemplateCrossDeptToMonth.xlsx";
		List<ExcelTable> tbList = new ArrayList<>();
		int tableNo = 1;
		List<CrossReportGroupModel> mockCrossReportTemplate = TestData.mockCrossReportTemplate(2, null);
		ExcelTable tp = new ExcelTable(tableNo, null, CrossReportGroupModel.class, Arrays.asList(mockCrossReportTemplate.get(0)));
		tp.setFirstRowIndex(6);
		tp.setNeedHead(false);
		tp.setTableName("第一个表");
		tp.setWriteBeforeHandler(new CrossTeamWriteBeforeHandlerImpl());
		tp.setDictionaryRefHandler(new AbstractDictionaryData());
//        tp.setTableDataDefaultFormat(DataFormatEnum.NUMBER_2_CURRENCY_US);
		tbList.add(tp);
		tableNo++;
		CrossReportGroupModel2 crossReportGroupModel2 = TestData.mockCrossReportTemplate();
		ExcelTable tp1 = new ExcelTable(tableNo, null, CrossReportGroupModel2.class, Arrays.asList(crossReportGroupModel2));
		tp1.setNeedHead(false);
		tbList.add(tp1);
		tableNo++;
//        List<CrossReportGroupModel> testTemplateCrossReportTestMode2 = TestData.emptyDataModel();
		List<CrossReportGroupModel> mockCrossReportTemplate1 = TestData.mockCrossReportTemplate(1, "market");
		ExcelTable tp2 = new ExcelTable(tableNo, null, CrossReportGroupModel.class, mockCrossReportTemplate1);
		tp2.setWriteBeforeHandler(new CrossMarketingWriteBeforeHandlerImpl());
		tp2.setNeedHead(false);
		tbList.add(tp2);
		tableNo++;
		Map<String, List<CrossReportModel>> mockDCCCrossReportData = TestData.mockDCCCrossReportData(15);
		int startRow = 10;
		for (Map.Entry<String, List<CrossReportModel>> map : mockDCCCrossReportData.entrySet()) {
			String key = map.getKey();
			List<CrossReportGroupNameModel> groupList = Arrays.asList(TestData.mockGroupCrossReportRate(key));
			ExcelTable group1 = new ExcelTable(tableNo, null, CrossReportGroupNameModel.class, groupList);
			ExcelShiftRange gShiftRange1 = new ExcelShiftRange();
			gShiftRange1.setStartRow(startRow);
			gShiftRange1.setShiftNumber(groupList.size());
			group1.setExcelShiftRange(gShiftRange1);
			group1.setEffectExcelFormula(false);
            ExcelStyle excelStyle = new ExcelStyle();
            ExcelFont excelTableFont = new ExcelFont();
            excelTableFont.setFontHeightInPoints((short) 11);
            excelTableFont.setBold(true);
            excelTableFont.setFontColor(IndexedColors.BLUE);
            excelStyle.setExcelFont(excelTableFont);
            excelStyle.setExcelBackGroundColor(IndexedColors.GREY_25_PERCENT);
            group1.setTableStyle(excelStyle);
			tableNo++;
			startRow += groupList.size();
			tbList.add(group1);
			List<CrossReportModel> value = map.getValue();
			ExcelTable excelTable1 = new ExcelTable(tableNo, null, CrossReportModel.class, value);
			excelTable1.setNeedHead(false);
//            excelTable1.setTableDataDefaultFormat(DataFormatEnum.NUMBER_2_CURRENCY_US);
			ExcelShiftRange shiftRange1 = new ExcelShiftRange();
			shiftRange1.setStartRow(startRow);
			shiftRange1.setShiftNumber(value.size());
			excelTable1.setExcelShiftRange(shiftRange1);
			excelTable1.setWriteBeforeHandler(new CrossDccWriteBeforeHandlerImpl());
			excelTable1.setDictionaryRefHandler(new AbstractDictionaryData2());
			startRow += value.size();
			tableNo++;
			tbList.add(excelTable1);
		}
		startRow++;
		if (!CollectionUtils.isEmpty(tbList)) {
			tableNo = tbList.get(tbList.size() - 1).getTableNo() + 1;
		}

		Map<String, List<CrossReportModel>> mockSCCrossReportData = TestData.mockSCCrossReportData(30);
		List<ExcelTable> firstScTables = new ArrayList<>();
		for (Map.Entry<String, List<CrossReportModel>> map : mockSCCrossReportData.entrySet()) {
			String key = map.getKey();
			List<CrossReportGroupNameModel> groupList = Arrays.asList(TestData.mockGroupCrossReportRate(key));
			ExcelTable group2 = new ExcelTable(tableNo, null, CrossReportGroupNameModel.class, groupList);
			group2.setNeedHead(false);
			ExcelShiftRange gShiftRange2 = new ExcelShiftRange();
			gShiftRange2.setStartRow(startRow);
			gShiftRange2.setShiftNumber(groupList.size());
			group2.setExcelShiftRange(gShiftRange2);
			group2.setEffectExcelFormula(false);
			ExcelStyle excelStyle = new ExcelStyle();
			ExcelFont excelTableFont = new ExcelFont();
			excelTableFont.setFontHeightInPoints((short) 11);
			excelTableFont.setBold(false);
			excelTableFont.setFontColor(IndexedColors.BLUE);
			excelStyle.setExcelFont(excelTableFont);
			excelStyle.setExcelBackGroundColor(IndexedColors.GREY_25_PERCENT);
			group2.setTableStyle(excelStyle);
			tableNo++;
			startRow += groupList.size();
			firstScTables.add(group2);
			List<CrossReportModel> value = map.getValue();
			ExcelTable excelTable2 = new ExcelTable(tableNo, null, CrossReportModel.class, value);
			excelTable2.setNeedHead(false);
			ExcelShiftRange shiftRange2 = new ExcelShiftRange();
			shiftRange2.setStartRow(startRow);
			shiftRange2.setShiftNumber(value.size());
			excelTable2.setExcelShiftRange(shiftRange2);
			excelTable2.setWriteBeforeHandler(new CrossSCWriteBeforeHandlerImpl());
			excelTable2.setWriteAfterHandler(new WriteAfterHandlerImpl());
			tableNo++;
			startRow += value.size();
			firstScTables.add(excelTable2);
		}
		tbList.addAll(firstScTables);
		int divideFormulaTRefTable = firstScTables.get(1).getTableNo();
		ExcelTable excelTable = firstScTables.get(firstScTables.size() - 1);
		int tableNo1 = excelTable.getTableNo();
		ExcelShiftRange excelShiftRange = excelTable.getExcelShiftRange();
		int initIndex = excelShiftRange.getStartRow() + excelShiftRange.getShiftNumber();
		System.out.println("startRow----:" + startRow);
		initIndex += 6;
		tableNo1++;
		Map<String, List<CrossReportRateModel>> tempCrossReportRateTestModel = TestData.mockSCCrossReportRateData(30);
		for (Map.Entry<String, List<CrossReportRateModel>> map : tempCrossReportRateTestModel.entrySet()) {
			String key = map.getKey();
			List<CrossReportRateGroupNameModel> groupList = Arrays.asList(TestData.mockGroupSCCrossReportRate(key));
			ExcelTable group2 = new ExcelTable(tableNo1, null, CrossReportRateGroupNameModel.class, groupList);
			group2.setNeedHead(false);
			ExcelShiftRange gShiftRange2 = new ExcelShiftRange();
			gShiftRange2.setStartRow(initIndex);
			gShiftRange2.setShiftNumber(groupList.size());
			group2.setExcelShiftRange(gShiftRange2);
			initIndex++;
			tableNo1++;
			tbList.add(group2);
			ExcelTable excelTable1 = new ExcelTable(tableNo1, null, CrossReportRateModel.class, map.getValue());
			excelTable1.setNeedHead(false);
			ExcelShiftRange gShiftRange3 = new ExcelShiftRange();
			gShiftRange3.setStartRow(initIndex);
			gShiftRange3.setShiftNumber(map.getValue().size());
			excelTable1.setExcelShiftRange(gShiftRange3);
			excelTable1.setTableDataDefaultFormat(DataFormatEnum.NUMBER_2_PERCENT);
			excelTable1.setDivideFormulaTRefTable(divideFormulaTRefTable);
			excelTable1.setWriteBeforeHandler(new CrossRateWriteBeforeHandlerImpl());
			tbList.add(excelTable1);
			initIndex += map.getValue().size();
			tableNo1++;
			divideFormulaTRefTable += 2;
		}
		ExcelSheet sheet1 = new ExcelSheet(0, tbList);
		Map<String, String> refMap = new HashMap<>();
		refMap.put("YYYY", LocalDate.now().getYear() + "");
		refMap.put("MM", LocalDate.now().getMonthValue() + "");
		sheet1.setWriteLoadTemplateHandler(new CrossWriteLoadTemplateHandlerImpl(refMap));
		sheet1.setSheetDataDefaultFormat(DataFormatEnum.NUMBER_2_ROUND_UP);
		ExcelFreezePaneRange freezePaneRange = new ExcelFreezePaneRange();
		freezePaneRange.setCellIndex(1);
		freezePaneRange.setCellCount(1);
		freezePaneRange.setRowCount(6);
		freezePaneRange.setRowIndex(7);
		sheet1.setFreezePaneRange(freezePaneRange);
		sheet1.setLocked(false);
		sheet1.setDefaultBackGroundColor(IndexedColors.GREY_25_PERCENT);
		List<ExcelSheet> sheets = new ArrayList<ExcelSheet>() {{
			add(sheet1);
		}};
		ExcelWriteParam excelWriteParam = new ExcelWriteParam();
//        excelWriteParam.setExcelFileName(outFileName);
//        excelWriteParam.setExcelOutFilePath(outFilePath);
		excelWriteParam.setExcelTemplateFile(outTemplate);
		excelWriteParam.setExcelOutFileFullPath(excelOutFileFullPath);
		excelWriteParam.setExcelSheets(sheets);
//        excelWriteParam.setEncryptPwd("lzg");
		ExcelWriteResponse export = IRISExcelFileFactory.exportV2007WithTemplate(excelWriteParam, resourcesFileInputStream);
		System.out.println("------------------------export result start-----------------");
		System.out.println(JSONUtil.objectToString(export));
		System.out.println("------------------------export result end-----------------");
		Assert.assertEquals(export.getCode(), FileConstant.SUCCESS_CODE);
	}

	@Test
	public void export2007WithQueueTemplate() {
		List<ExcelTable> tableList = new ArrayList<>();
		for (int i = 1; i <= 100; i++) {
			List<TemplateModel> list = TestData.mockTemplateModel(20);
			ExcelTable tp = new ExcelTable(i, null, TemplateModel.class, list);
			if (i == 1) {
				tp.setFirstRowIndex(1);
			}
			tp.setNeedHead(false);
			tp.setDictionaryRefHandler(new TemplateDic());
			tableList.add(tp);
		}
		String outFilePath = "D:\\导出EXCEL";
		String outFileName = ExcelV2007Util.generateUniqueFileNameByTime("export") + ".xlsx";
		String outTemplate = "template.xlsx";
//        String excelOutFileFullPath = outFilePath + "\\" + outFileName;
		ExcelSheet sheet1 = new ExcelSheet(0, tableList);
		sheet1.setSheetDataDefaultFormat(DataFormatEnum.NUMBER_2_ROUND_UP);
		ExcelFreezePaneRange freezePaneRange = new ExcelFreezePaneRange();
		freezePaneRange.setRowCount(1);
		freezePaneRange.setRowIndex(1);
		sheet1.setFreezePaneRange(freezePaneRange);
		sheet1.setDefaultBackGroundColor(IndexedColors.LIGHT_BLUE);
//		sheet1.setDefaultBackGroundColor(IndexedColors.GREY_40_PERCENT);
		List<ExcelSheet> sheets = new ArrayList<ExcelSheet>() {{
			add(sheet1);
		}};
		ExcelWriteParam excelWriteParam = new ExcelWriteParam();
		excelWriteParam.setExcelFileName(outFileName);
		excelWriteParam.setExcelOutFilePath(outFilePath);
		excelWriteParam.setExcelTemplateFile(outTemplate);
//        excelWriteParam.setExcelOutFileFullPath(excelOutFileFullPath);
		excelWriteParam.setExcelSheets(sheets);
		InputStream resourcesFileInputStream = FileUtil.getResourcesFileInputStream(outTemplate);
		ExcelWriteResponse export = IRISExcelFileFactory.exportV2007QueueWithTemplate(excelWriteParam, resourcesFileInputStream);
//      ExcelWriteResponse export = IRISExcelFileFactory.exportV2007WithTemplate(excelWriteParam, resourcesFileInputStream);
		System.out.println("------------------------export result start-----------------");
		System.out.println(JSONUtil.objectToString(export));
		System.out.println("------------------------export result end-----------------");
		Assert.assertEquals(export.getCode(), FileConstant.SUCCESS_CODE);

	}


	@Test
	public void encryptExcel() throws Exception {
		String filePath = "D:\\导出模板\\tmp.xlsx";
		ExcelV2007Util.encrypt(filePath, "lzg");

	}

	@Test
	public void decryptExcel() throws Exception {
		String filePath = "D:\\导出模板\\tmp.xlsx";
		Workbook workbook = ExcelV2007Util.decrypt(filePath, "lzg");
		Sheet sheetAt = workbook.getSheetAt(0);
		Row row = sheetAt.getRow(6);
		Cell cell = row.getCell(2);
		System.out.println(cell.getNumericCellValue());
	}

}
