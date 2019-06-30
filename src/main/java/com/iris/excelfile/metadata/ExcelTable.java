package com.iris.excelfile.metadata;


import com.iris.excelfile.constant.DataFormatEnum;
import com.iris.excelfile.constant.TableLayoutEnum;
import com.iris.excelfile.core.handler.extend.DictionaryRefHandler;
import com.iris.excelfile.core.handler.extend.IWriteAfterHandler;
import com.iris.excelfile.core.handler.extend.IWriteBeforeHandler;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author liu_wp
 * @date Created in 2019/3/5 12:17
 * @see
 */
public class ExcelTable<T> implements Comparable<ExcelTable> {
	private int tableNo;
	private String tableName;
	private ExcelHeadProperty excelHeadProperty;
	private boolean needHead;
	private List<T> data;
	private ExcelStyle tableStyle;
	private ExcelStyle tableHeadStyle;
	private List<ExcelColumnProperty> columnPropertyList = new ArrayList<>();
	private Class<? extends BaseRowModel> headClass;
	private Map<Integer, CellStyle> cellStyleMap;
	/**
	 * 注解表头信息
	 */
	private List<List<String>> head = new ArrayList<>();
	/**
	 * 开始行
	 */
	private int firstRowIndex;

	/**
	 * 开始列
	 */
	private int firstCellIndex;
	/**
	 * 行高
	 */
	private Integer heightInPoints;
	/**
	 * table 之间的行数间隔
	 */
	private int spaceNum;
	/**
	 * table 布局方向
	 */
	private TableLayoutEnum tableLayoutEnum;

	/**
	 * 写入文档之后执行
	 */
	private IWriteAfterHandler writeAfterHandler;
	/**
	 * 写入文档之前执行
	 */
	private IWriteBeforeHandler writeBeforeHandler;
	/**
	 * 字典映射handler
	 */
	private DictionaryRefHandler dictionaryRefHandler;
	/**
	 * 模板移动范围 布局优先
	 */
	private ExcelShiftRange excelShiftRange;
	/**
	 * 表格范围
	 */
	private ExcelCellRange tableCellRange;
	/**
	 * 是否生效table 的计算公式
	 */
	private boolean isEffectExcelFormula = true;
	/**
	 * 除法计算公式 关联的table 获取计算开始的行数索引
	 */
	private Integer divideFormulaTRefTable;
	/**
	 * 默认表格的数字格式
	 */
	private DataFormatEnum tableDataDefaultFormat;
	/**
	 * 锁文档
	 */
	private boolean locked;
	/**
	 * 空数据的时候设置的默认行数 避免空数据布局错乱
	 */
	private int minDataRowCount;
	/**
	 * head + firstRowIndex
	 */
	private int startContentRowIndex = 0;

	public ExcelTable() {

	}

	public ExcelTable(int tableNo, List<List<String>> head, Class<? extends BaseRowModel> headClass, List<T> data) {
		this(tableNo, head, headClass, true, false, data, 0, 0, null, null, null);
	}

	public ExcelTable(int tableNo, List<List<String>> head, Class<? extends BaseRowModel> headClass, List<T> data, boolean needHead) {
		this(tableNo, head, headClass, needHead, false, data, 0, 0, null, null, null);
	}

	public ExcelTable(int tableNo, List<List<String>> head, Class<? extends BaseRowModel> headClass, List<T> data, int firstRowIndex, int firstCellIndex) {
		this(tableNo, head, headClass, true, false, data, 0, 0, null, null, null);
	}

	public ExcelTable(int tableNo, List<List<String>> head, Class<? extends BaseRowModel> headClass, boolean needHead, boolean locked, List<T> data, int firstRowIndex, int firstCellIndex, TableLayoutEnum tableLayoutEnum, IWriteBeforeHandler iWriteBeforeHandler, IWriteAfterHandler iWriteAfterHandler) {
		this.tableNo = tableNo <= 0 ? 1 : tableNo;
		this.tableName = "excelTable" + this.tableNo;
		this.head = head;
		this.locked = locked;
		this.headClass = headClass;
		this.data = data;
		this.needHead = !needHead ? needHead : true;
		this.firstRowIndex = firstRowIndex >= 0 ? firstRowIndex : 0;
		this.firstCellIndex = firstCellIndex >= 0 ? firstCellIndex : 0;
		this.writeBeforeHandler = iWriteBeforeHandler;
		this.writeAfterHandler = iWriteAfterHandler;
		this.tableLayoutEnum = tableLayoutEnum != null ? tableLayoutEnum : TableLayoutEnum.BOTTOM;
	}


	public int getTableNo() {
		return tableNo;
	}

	public void setTableNo(int tableNo) {
		this.tableNo = tableNo <= 0 ? 1 : tableNo;
	}

	public ExcelHeadProperty getExcelHeadProperty() {
		return excelHeadProperty;
	}

	public void setExcelHeadProperty(ExcelHeadProperty excelHeadProperty) {
		this.excelHeadProperty = excelHeadProperty;
	}

	public ExcelCellRange getTableCellRange() {
		return tableCellRange;
	}

	public void setTableCellRange(ExcelCellRange tableCellRange) {
		this.tableCellRange = tableCellRange;
	}

	public ExcelStyle getTableStyle() {
		return tableStyle;
	}

	public void setTableStyle(ExcelStyle tableStyle) {
		this.tableStyle = tableStyle;
	}

	public List<ExcelColumnProperty> getColumnPropertyList() {
		return columnPropertyList;
	}

	public void setColumnPropertyList(List<ExcelColumnProperty> columnPropertyList) {
		this.columnPropertyList = columnPropertyList;
	}

	public IWriteBeforeHandler getWriteBeforeHandler() {
		return writeBeforeHandler;
	}

	public void setWriteBeforeHandler(IWriteBeforeHandler writeBeforeHandler) {
		this.writeBeforeHandler = writeBeforeHandler;
	}

	public List<List<String>> getHead() {
		return head;
	}

	public void setHead(List<List<String>> head) {
		this.head = head;
	}

	public Class<? extends BaseRowModel> getHeadClass() {
		return headClass;
	}

	public void setHeadClass(Class<? extends BaseRowModel> headClass) {
		this.headClass = headClass;
	}

	public int getFirstRowIndex() {
		return firstRowIndex;
	}

	public void setFirstRowIndex(int firstRowIndex) {
		this.firstRowIndex = firstRowIndex;
	}

	public int getFirstCellIndex() {
		return firstCellIndex;
	}

	public void setFirstCellIndex(int firstCellIndex) {
		this.firstCellIndex = firstCellIndex;
	}

	public int getStartContentRowIndex() {
		return startContentRowIndex;
	}

	public void setStartContentRowIndex(int startContentRowIndex) {
		this.startContentRowIndex = startContentRowIndex;
	}

	public ExcelStyle getTableHeadStyle() {
		return tableHeadStyle;
	}

	public void setTableHeadStyle(ExcelStyle tableHeadStyle) {
		this.tableHeadStyle = tableHeadStyle;
		if (this.excelHeadProperty != null) {
			this.excelHeadProperty.setHeadStyle(tableHeadStyle);
		}
	}

	public int getSpaceNum() {
		return spaceNum;
	}

	public void setSpaceNum(int spaceNum) {
		this.spaceNum = spaceNum >= 0 ? spaceNum : 0;
	}

	public TableLayoutEnum getTableLayoutEnum() {
		return tableLayoutEnum;
	}

	public void setTableLayoutEnum(TableLayoutEnum tableLayoutEnum) {
		if (tableLayoutEnum == null) {
			this.tableLayoutEnum = TableLayoutEnum.BOTTOM;
		} else {
			this.tableLayoutEnum = tableLayoutEnum;
		}
	}

	public boolean isNeedHead() {
		return needHead;
	}

	public void setNeedHead(boolean needHead) {
		this.needHead = needHead;
	}

	public IWriteAfterHandler getWriteAfterHandler() {
		return writeAfterHandler;
	}

	public void setWriteAfterHandler(IWriteAfterHandler writeAfterHandler) {
		this.writeAfterHandler = writeAfterHandler;
	}

	public boolean isLocked() {
		return locked;
	}

	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	public boolean isEffectExcelFormula() {
		return isEffectExcelFormula;
	}

	public void setEffectExcelFormula(boolean effectExcelFormula) {
		isEffectExcelFormula = effectExcelFormula;
	}

	public ExcelShiftRange getExcelShiftRange() {
		return excelShiftRange;
	}

	public void setExcelShiftRange(ExcelShiftRange excelShiftRange) {
		this.excelShiftRange = excelShiftRange;
	}

	public Integer getDivideFormulaTRefTable() {
		return divideFormulaTRefTable;
	}

	public void setDivideFormulaTRefTable(Integer divideFormulaTRefTable) {
		this.divideFormulaTRefTable = divideFormulaTRefTable;
	}

	public String getTableName() {
		return tableName;
	}

	public void setTableName(String tableName) {
		this.tableName = tableName;
	}

	public int getMinDataRowCount() {
		return minDataRowCount;
	}

	public void setMinDataRowCount(int minDataRowCount) {
		this.minDataRowCount = minDataRowCount;
	}

	public DataFormatEnum getTableDataDefaultFormat() {
		return tableDataDefaultFormat;
	}

	public void setTableDataDefaultFormat(DataFormatEnum tableDataDefaultFormat) {
		this.tableDataDefaultFormat = tableDataDefaultFormat;
	}

	public Integer getHeightInPoints() {
		return heightInPoints;
	}

	public void setHeightInPoints(Integer heightInPoints) {
		this.heightInPoints = heightInPoints;
	}

	public DictionaryRefHandler getDictionaryRefHandler() {
		return dictionaryRefHandler;
	}

	public void setDictionaryRefHandler(DictionaryRefHandler dictionaryRefHandler) {
		this.dictionaryRefHandler = dictionaryRefHandler;
	}

	public Map<Integer, CellStyle> getCellStyleMap() {
		return cellStyleMap;
	}

	public void setCellStyleMap(Map<Integer, CellStyle> cellStyleMap) {
		this.cellStyleMap = cellStyleMap;
	}

	public List<T> getData() {
		return data;
	}

	public void setData(List<T> data) {
		this.data = data;
	}

	@Override
	public int compareTo(ExcelTable o) {
		int x = this.tableNo;
		int y = o.getTableNo();
		return (x < y) ? -1 : ((x == y) ? 0 : 1);
	}
}
