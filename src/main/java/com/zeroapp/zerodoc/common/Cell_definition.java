package com.zeroapp.zerodoc.common;

import java.util.ArrayList;
import java.util.List;

import org.dom4j.Element;

/**
 * 每个cell的定义，每个Cell中可以包含多个Cell_data
 * 其中row和column都为相对坐标，即相对于所在range左上角的坐标，使用时需要转换为绝对坐标
 * 
 * @author wkm
 * 
 */
public class Cell_definition
{
	private int _row;
	private int _column;
	private List<Data_unit> _cell_datas = null;
	private Element _style_info; // 样式，xml格式

	public Element get_style_info() {
		return _style_info;
	}

	public void set_style_info(Element _style_info) {
		this._style_info = _style_info;
	}

	/**
	 * 构造函数，为某个cell添加一个cell_data列表，复杂情况下使用
	 * 
	 * @param row
	 * @param column
	 * @param cell_datas
	 * @param style_info
	 */
	public Cell_definition(int row, int column, List<Data_unit> cell_datas, Element style_info)
	{
		_row = row;
		_column = column;
		_cell_datas = cell_datas;
		_style_info = style_info;
	}

	/**
	 * 构造函数，为某个cell添加一个cell_data
	 * 
	 * @param row
	 * @param column
	 * @param cell_data
	 * @param style_info
	 */
	public Cell_definition(int row, int column, Data_unit cell_data, Element style_info)
	{
		_row = row;
		_column = column;
		_cell_datas = new ArrayList();
		_cell_datas.add(cell_data);
		_style_info = style_info;
	}

	public int get_row()
	{
		return _row;
	}

	public void set_row(int _row)
	{
		this._row = _row;
	}

	public int get_column()
	{
		return _column;
	}

	public void set_column(int _column)
	{
		this._column = _column;
	}

	public List<Data_unit> get_cell_datas()
	{
		return _cell_datas;
	}

	public void set_cell_datas(List<Data_unit> _datas)
	{
		this._cell_datas = _datas;
	}

}
