package com.zeroapp.zerodoc.msword;

import java.util.ArrayList;
import java.util.List;

import org.dom4j.Element;

import com.jacob2.com.Dispatch;
import com.jacob2.com.Variant;
import com.zeroapp.zerodoc.exceptions.ZerodocException;
import com.zeroapp.zerodoc.msword.util.BorderStyle;
import com.zeroapp.zerodoc.msword.util.ConstantUtil;

public class Table {
	Dispatch _table = null;

	/**
	 * 构造函数
	 * 
	 * @param table
	 */
	public Table(Variant table) {
		_table = table.getDispatch();
	}

	/**
	 * 返回表格中的某个单元格
	 * 
	 * @param row
	 *            行数，从1开始
	 * @param column
	 *            列数，从1开始
	 */
	public Cell getCell(int row, int column) {
		Dispatch cell = Dispatch.call(_table, "Cell", new Variant(row),
				new Variant(column)).toDispatch();
		return new Cell(cell);

	}

	/**
	 * 清除表格的样式设定
	 */
	public void clearFormatting() {
		this.select();
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// 清除格式
		Dispatch.call(selection, "ClearFormatting");
	}

	/**
	 * 合并表格中的两个单元格，以及两者之间的其他单元格
	 * 注意，当前面有合并的单元格时，定位需要重新计算
	 * 
	 * @param row1
	 *            行数，从1开始
	 * @param column1
	 *            列数，从1开始
	 * @param row2
	 *            行数，从1开始
	 * @param column2
	 *            列数，从1开始
	 * @throws ZerodocException 
	 */
	public void mergeCells(int row1, int column1, int row2, int column2) throws ZerodocException {
		Dispatch cell_1 = Dispatch.call(_table, "Cell", new Variant(row1),
				new Variant(column1)).toDispatch();
		Dispatch cell_2 = Dispatch.call(_table, "Cell", new Variant(row2),
				new Variant(column2)).toDispatch();
		try {
			Dispatch.callSub(cell_1, "Merge", cell_2);
		} catch (Exception e) {
			throw new ZerodocException("无法合并单元格[" + row1 + "," + column1
					+ "] to cell[" + row2 + "," + column2 + "]");
		}
	}

	/**
	 * 批量合并，可以解决由于前面单元格合并后造成的定位不准问题
	 * 
	 * @param merge_info
	 * @throws ZerodocException 
	 */
	public void batchMerge(Element merge_infos) throws ZerodocException {
		if (merge_infos == null) {
			return;
		}

		try{
		List<Dispatch> beginCells = new ArrayList();
		List<Dispatch> endCells = new ArrayList();

		List<Element> merge_info_list = merge_infos.elements("merge_info");

		// 首先将merge信息全部保存
		for (Element merge_info : merge_info_list) {
			String merge_start_row = merge_info.attributeValue("start_row");
			String merge_start_column = merge_info
					.attributeValue("start_column");
			String merge_end_row = merge_info.attributeValue("end_row");
			String merge_end_column = merge_info.attributeValue("end_column");

			Dispatch beginCell = Dispatch.call(_table, "Cell",
					new Variant(Integer.parseInt(merge_start_row)),
					new Variant(Integer.parseInt(merge_start_column)))
					.toDispatch();

			Dispatch endCell = Dispatch.call(_table, "Cell",
					new Variant(Integer.parseInt(merge_end_row)),
					new Variant(Integer.parseInt(merge_end_column)))
					.toDispatch();

			beginCells.add(beginCell);
			endCells.add(endCell);
		}
		
		for(int i = 0; i < beginCells.size(); i++){
			Dispatch.callSub(beginCells.get(i), "Merge", endCells.get(i));
		}
		}catch(Exception e){
			throw new ZerodocException("合并单元格失败：" + e.getMessage());			
		}

	}

	/**
	 * 添加新行
	 * 
	 * @param count
	 */
	public void addRow(int count) {
		if (count > 0) {
			Dispatch rows = Dispatch.get(_table, "Rows").toDispatch();
			for (int i = 0; i < count; i++) {
				Dispatch.call(rows, "Add");
			}
		}
	}

	/**
	 * 添加新列
	 * 
	 * @param count
	 */
	public void addColumn(int count) {
		if (count > 0) {
			Dispatch columns = Dispatch.get(_table, "Columns").toDispatch();
			for (int i = 0; i < count; i++) {
				Dispatch.call(columns, "Add");
			}
		}
	}

	/**
	 * 返回当前表格的行数
	 * 
	 * @return
	 */
	public int getRowCount() {
		Dispatch columns = Dispatch.get(_table, "Rows").toDispatch();
		return Dispatch.get(columns, "Count").getInt();

	}

	/**
	 * 返回当前表格的列数
	 * 
	 * @return
	 */
	public int getColumnCount() {
		Dispatch columns = Dispatch.get(_table, "Columns").toDispatch();
		return Dispatch.get(columns, "Count").getInt();

	}

	/**
	 * 选中表格
	 */
	public void select() {
		Dispatch.callSub(_table, "Select");
	}

	/**
	 * 设定表格的对齐方式
	 * 
	 * @param alignment
	 */
	public void setAlignment(int alignment) {
		this.select();
		// 得到selection对象
		Dispatch rows = Dispatch.get(_table, "Rows").toDispatch();
		Dispatch.put(rows, "Alignment", new Variant(alignment));
	}

	/**
	 * 设定表格自动调整宽度的方式 包括按内容调整、按窗口调整、固定宽度三种类型
	 * 
	 * @param behavior
	 */
	public void setAutoFitBehavior(int behavior) {
		Dispatch.callSub(_table, "AutoFitBehavior", new Variant(behavior));
	}

	/**
	 * 设定表格的边框
	 * 
	 * @param borderStyle
	 */
	public void setBorderStyle(BorderStyle borderStyle) {
		if (borderStyle == null) {
			return;
		}
		// 得到borders对象
		Dispatch borders = Dispatch.get(_table, "Borders").toDispatch();
		if (borderStyle.getBorderTopStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderTop)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderTopStyle()));
		}
		if (borderStyle.getBorderBottompStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderBottom)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderBottompStyle()));
		}
		if (borderStyle.getBorderLeftStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderLeft)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderLeftStyle()));
		}
		if (borderStyle.getBorderRightStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderRight)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderRightStyle()));
		}
		if (borderStyle.getBorderHorizontalStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderHorizontal)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderHorizontalStyle()));
		}
		if (borderStyle.getBorderVerticalStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderVertical)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderVerticalStyle()));
		}
	}

	/**
	 * 设定某个区域的边框样式
	 * 
	 * @param borderStyle
	 * @param startRow
	 * @param startColumn
	 * @param endRow
	 * @param endColumn
	 * @throws Exception
	 */
	public void SetRangeBorderStyle(BorderStyle borderStyle, int startRow,
			int startColumn, int endRow, int endColumn) throws Exception {

		if (startRow < 1 || startColumn < 1 || endRow < 1 || endColumn < 1) {
			throw new Exception("Range的行列设置非法，必须大于0");
		}
		// 调整起点和终点
		if (startRow > endRow) {
			int temp = startRow;
			startRow = endRow;
			endRow = temp;
		}
		if (startColumn > endColumn) {
			int temp = startColumn;
			startColumn = endColumn;
			endColumn = temp;
		}

		this.getCell(startRow, startColumn).select();
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// 扩展Selection对象
		Dispatch.call(selection, "MoveDown", new Variant(ConstantUtil.wdLine),
				new Variant(endRow - startRow), new Variant(
						ConstantUtil.wdExtend));
		Dispatch.call(selection, "MoveRight", new Variant(
				ConstantUtil.wdCharacter),
				new Variant(endColumn - startColumn), new Variant(
						ConstantUtil.wdExtend));

		// 得到borders对象
		Dispatch borders = Dispatch.get(selection, "Borders").toDispatch();
		if (borderStyle.getBorderTopStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderTop)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderTopStyle()));
		}
		if (borderStyle.getBorderBottompStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderBottom)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderBottompStyle()));
		}
		if (borderStyle.getBorderLeftStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderLeft)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderLeftStyle()));
		}
		if (borderStyle.getBorderRightStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderRight)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderRightStyle()));
		}
		if (borderStyle.getBorderHorizontalStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderHorizontal)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderHorizontalStyle()));
		}
		if (borderStyle.getBorderVerticalStyle() != null) {
			Dispatch border = Dispatch.call(borders, "Item",
					new Variant(BorderStyle.wdBorderVertical)).toDispatch();
			Dispatch.put(border, "LineStyle", new Variant(borderStyle
					.getBorderVerticalStyle()));
		}

	}
}
