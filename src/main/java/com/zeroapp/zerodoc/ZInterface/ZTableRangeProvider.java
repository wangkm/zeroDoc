package com.zeroapp.zerodoc.ZInterface;

import com.zeroapp.zerodoc.common.TableRange_data;

public interface ZTableRangeProvider extends ZProvider
{
	/**
	 * 回调函数，由调用者设定当前表格区域的起始行和起始列
	 * @param start_row
	 */
	public void set_start_row(int start_row);
	
	public void set_start_column(int start_column);

	public void set_end_row(int end_row);
	
	public void set_end_column(int end_column);

	/**
	 * 返回表格数据
	 * 
	 * @return
	 */
	public TableRange_data get_range_data();

}
