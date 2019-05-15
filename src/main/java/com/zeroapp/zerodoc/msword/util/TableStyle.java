package com.zeroapp.zerodoc.msword.util;

public class TableStyle
{
	private String style = null;
	private boolean ApplyStyleHeadingRows = true;
	private boolean ApplyStyleLastRow = true;
	private boolean ApplyStyleFirstColumn = true;
	private boolean ApplyStyleLastColumn = true;
	
	public String getStyle()
	{
		return style;
	}
	public void setStyle(String style)
	{
		this.style = style;
	}
	public boolean isApplyStyleHeadingRows()
	{
		return ApplyStyleHeadingRows;
	}
	public void setApplyStyleHeadingRows(boolean applyStyleHeadingRows)
	{
		ApplyStyleHeadingRows = applyStyleHeadingRows;
	}
	public boolean isApplyStyleLastRow()
	{
		return ApplyStyleLastRow;
	}
	public void setApplyStyleLastRow(boolean applyStyleLastRow)
	{
		ApplyStyleLastRow = applyStyleLastRow;
	}
	public boolean isApplyStyleFirstColumn()
	{
		return ApplyStyleFirstColumn;
	}
	public void setApplyStyleFirstColumn(boolean applyStyleFirstColumn)
	{
		ApplyStyleFirstColumn = applyStyleFirstColumn;
	}
	public boolean isApplyStyleLastColumn()
	{
		return ApplyStyleLastColumn;
	}
	public void setApplyStyleLastColumn(boolean applyStyleLastColumn)
	{
		ApplyStyleLastColumn = applyStyleLastColumn;
	}

}
