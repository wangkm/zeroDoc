package com.zeroapp.zerodoc.common;

import java.util.List;

/**
 * Table区域的数据
 * 
 * @author wkm
 * 
 */
public class TableRange_data
{
	// cell列表
	private List<Cell_definition> _cell_infos = null;
	private List<Merge_definition> _merge_infos = null;

	public TableRange_data(List<Cell_definition> cell_infos,
			List<Merge_definition> merge_infos)
	{
		_cell_infos = cell_infos;
		_merge_infos = merge_infos;
	}

	public List<Cell_definition> get_cell_infos()
	{
		return _cell_infos;
	}

	public void set_cell_infos(List<Cell_definition> _cell_infos)
	{
		this._cell_infos = _cell_infos;
	}

	public List<Merge_definition> get_merge_infos()
	{
		return _merge_infos;
	}

	public void set_merge_infos(List<Merge_definition> _merge_infos)
	{
		this._merge_infos = _merge_infos;
	}

}
