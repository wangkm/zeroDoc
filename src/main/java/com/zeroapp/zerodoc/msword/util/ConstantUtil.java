package com.zeroapp.zerodoc.msword.util;

public class ConstantUtil
{
	// 几个枚举变量
	static public int wdPageBreak = 7;
	
	// 节的类型
	static public int wdSectionBreakNextPage = 2;
	static public int wdSectionBreakContinuous = 3;

	static public int wdSeekMainDocument = 0;
	static public int wdSeekCurrentPageHeader = 9;
	static public int wdSeekCurrentPageFooter = 10;

	// cell的高度设定属性
	static public int wdRowHeightAuto = 0;
	static public int wdRowHeightAtLeast = 1;
	static public int wdRowHeightExactly = 2;

	// 页码等
	static public int wdFieldNumPages = 26;
	static public int wdFieldPage = 33;
	static public int wdFieldDate = 31;
	static public int wdFieldTime = 32;
	
	// 缺省页边距
	static final public double default_topMargin = 2.54;
	static final public double default_buttomMargin = 2.54;
	static final public double default_leftMargin = 3.17;
	static final public double default_rightMargin = 3.17;
	
	// 缺省页眉和页脚的距离
	static final public double default_headerDistance = 1.5;
	static final public double default_footerDistance = 1.75;
	
	static final public int wdCharacter = 1;
	static final public int wdLine = 5;
	static final public int wdExtend = 1;
	static final public int wdMove = 0;
	
	// 关闭word实例时是否保存文档的标志位
	static final public int wdDoNotSaveChanges = 0;
	static final public int wdSaveChanges = -1;
	static final public int wdPromptToSaveChanges = -2;

	// 页眉页脚的类型
	static final public int wdHeaderFooterPrimary = 1;
	static final public int wdHeaderFooterFirstPage = 2;
	static final public int wdHeaderFooterEvenPages = 3;
		
	// word的内建属性
	static final public int wdPropertyPages = 14;		// 总页数
	
}
