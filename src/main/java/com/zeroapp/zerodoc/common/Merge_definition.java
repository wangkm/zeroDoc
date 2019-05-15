package com.zeroapp.zerodoc.common;

/**
 * 需要合并的cell定义
 * _start_row, _start_column, 左上角单元格的位置
 * _end_row, _end_column, 右下角单元格的位置
 * 
 * @author wkm
 * 
 */
public class Merge_definition
{
	private int _start_row;
	private int _start_column;
	private int _end_row;
	private int _end_column;

	public Merge_definition(int start_row, int start_column, int end_row,
			int end_column)
	{
		_start_row = start_row;
		_start_column = start_column;
		_end_row = end_row;
		_end_column = end_column;
	}

	public int get_start_row()
	{
		return _start_row;
	}

	public void set_start_row(int _start_row)
	{
		this._start_row = _start_row;
	}

	public int get_start_column()
	{
		return _start_column;
	}

	public void set_start_column(int _start_column)
	{
		this._start_column = _start_column;
	}

	public int get_end_row()
	{
		return _end_row;
	}

	public void set_end_row(int _end_row)
	{
		this._end_row = _end_row;
	}

	public int get_end_column()
	{
		return _end_column;
	}

	public void set_end_column(int _end_column)
	{
		this._end_column = _end_column;
	}

}
