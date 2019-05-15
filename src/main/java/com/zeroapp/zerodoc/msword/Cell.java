package com.zeroapp.zerodoc.msword;

import org.dom4j.Element;

import com.jacob2.activeX.ActiveXComponent;
import com.jacob2.com.Dispatch;
import com.jacob2.com.Variant;
import com.zeroapp.zerodoc.exceptions.ZerodocException;
import com.zeroapp.zerodoc.msword.util.BorderStyle;
import com.zeroapp.zerodoc.msword.util.ConstantUtil;
import com.zeroapp.zerodoc.msword.util.FontStyle;
import com.zeroapp.zerodoc.msword.util.ParaStyle;
import com.zeroapp.zerodoc.msword.util.VerticalAlignmentType;
import com.zeroapp.zerodoc.util.MiscUtil;

public class Cell {
	Dispatch _cell = null;

	public Cell(Variant cell) {
		_cell = cell.getDispatch();
	}

	public Cell(Dispatch cell) {
		_cell = cell;
	}

	/**
	 * 设定某个单元格所在列的宽度
	 * 
	 * @param width
	 * @throws ZerodocException
	 */
	private void setWidthOfColumn(float width) throws ZerodocException {
		// Dispatch.put(_cell, "Width", new Variant(width));
		try {
			Dispatch column = Dispatch.call(_cell, "Column").toDispatch();
			Dispatch.put(column, "Width", new Variant(width));
		} catch (Exception e) {
			throw new ZerodocException("调整单元格所在列的宽度失败：" + e.getMessage());
		}
	}

	/**
	 * 设定某个单元格所在行的高度
	 * 
	 * @param height
	 * @throws ZerodocException
	 */
	private void setHeightOfRow(float height) throws ZerodocException {
		try {
			Dispatch row = ActiveXComponent.call(_cell, "Row").toDispatch();
			Dispatch.put(row, "HeightRule", ConstantUtil.wdRowHeightAtLeast);
			Dispatch.put(row, "Height", new Variant(height));
		} catch (Exception e) {
			throw new ZerodocException("调整单元格所在列的高度失败：" + e.getMessage());
		}
	}

	/**
	 * 设定某个单元格的宽度
	 * 
	 * @param width
	 */
	private void setWidthSingle(float width) {
		Dispatch.put(_cell, "Width", new Variant(width));
	}

	/**
	 * 设定某个单元格的高度
	 * 
	 * @param height
	 */
	private void setHeightSingle(float height) {
		Dispatch.put(_cell, "HeightRule", ConstantUtil.wdRowHeightAtLeast);
		Dispatch.put(_cell, "Height", new Variant(height));
	}

	/**
	 * 清除单元格的样式设定
	 */
	public void clearFormatting() {
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// 清除格式
		Dispatch.call(selection, "ClearFormatting");
	}

	/**
	 * 设定当前单元格为选中状态
	 */
	public void select() {
		Dispatch.callSub(_cell, "Select");
	}
	
	/**
	 * 设定单元格的边框 注意，单元格的item中不存在wdBorderVertical和wdBorderHorizontal
	 * 
	 * @param borderStyle
	 */
	private void setBorderStyle(BorderStyle borderStyle) {
		if (borderStyle == null) {
			return;
		}

		// 得到borders对象
		Dispatch borders = Dispatch.get(_cell, "Borders").toDispatch();
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
	}

	/**
	 * 设定单元格的垂直对其方式
	 * 
	 * @param verticalAligment
	 * @throws Exception
	 */
	private void setVerticalAlignment(int verticalAligment) throws Exception {
		Dispatch.put(_cell, "VerticalAlignment", verticalAligment);
	}

	/**
	 * 设定cell的样式
	 * 
	 * @param style_info
	 * @throws Exception
	 */
	private void setHeightAndWidth(Element style_info) throws Exception {
		if (style_info == null) {
			return;
		}

		// 设定单元格的宽度和高度，单位为厘米
		String cell_width = style_info.attributeValue("width");
		if (cell_width != null) {
			try {
				this.setWidthOfColumn(Float.parseFloat(cell_width));
			} catch (Exception e) {
				try {
					System.out.println("未能设定单元格所在列的宽度：" + e.getMessage()
							+ ", 尝试单独设定当前单元格的宽度");
					this.setWidthSingle(Float.parseFloat(cell_width));
				} catch (Exception e2) {
					throw e2;
				}
			}
		}
		String cell_height = style_info.attributeValue("height");
		if (cell_height != null) {
			try {
				this.setHeightOfRow(Float.parseFloat(cell_height));
			} catch (Exception e) {
				try {
					System.out.println("未能设定单元格所在行的高度：" + e.getMessage()
							+ ", 尝试单独设定当前单元格的高度");
					this.setHeightSingle(Float.parseFloat(cell_height));
				} catch (Exception e2) {
					throw e2;
				}
			}
		}

	}
	

	/**
	 * 设定单元格的样式
	 * 
	 * @param style_info
	 * @throws Exception
	 */
	public void setStyle(Element style_info, boolean inherited)
			throws Exception {
		if (style_info == null) {
			return;
		}
		
		// 设定单元格高度和宽度
		this.setHeightAndWidth(style_info);

		// 设定单元格垂直对齐方式
		String vAligment = style_info.attributeValue("vAlignment");
		if (vAligment != null) {
			this.setVerticalAlignment(VerticalAlignmentType
					.getVerticalAlignmentID(vAligment));
		}

		// 得到当前cell内的Range对象
		Dispatch range = Dispatch.get(_cell, "Range").toDispatch();
		Range range_obj = new Range(range);
		
		// 设定单元格内部对象的样式
		range_obj.setStyle(style_info);
				
		
		// 如果是非继承样式设定的化，设定单元格边框
		if(inherited){
			BorderStyle borderStyle = new BorderStyle(style_info);
			this.setBorderStyle(borderStyle);
		}

	}

	/**
	 * 将光标移动到单元格的末尾
	 * 
	 * @throws ZerodocException
	 */
	public void moveCursorToEnd() throws ZerodocException {
		try {
			Dispatch range = Dispatch.get(_cell, "Range").toDispatch();
			int endKey = Dispatch.get(range, "End").toInt();
			Dispatch.call(range, "SetRange", endKey - 1, endKey - 1);
			Dispatch.call(range, "Select");
		} catch (Exception e) {
			throw new ZerodocException(e.getMessage());
		}
	}
	
	public Range getLastRange() throws Exception{
		this.moveCursorToEnd();
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
		.toDispatch();
		Dispatch range = Dispatch.get(selection, "Range")
		.toDispatch();
		return new Range(range);
		
	}
}
