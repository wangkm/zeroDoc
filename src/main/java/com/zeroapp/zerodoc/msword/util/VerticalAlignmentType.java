package com.zeroapp.zerodoc.msword.util;

import com.zeroapp.zerodoc.exceptions.ZerodocException;

public class VerticalAlignmentType
{
	public static int getVerticalAlignmentID(String valigment) throws Exception
	{
		if ("wdCellAlignVerticalTop".equalsIgnoreCase(valigment)){
			return 0;
		}else if("wdCellAlignVerticalCenter".equalsIgnoreCase(valigment)){
			return 1;
		}else if("wdCellAlignVerticalBottom".equalsIgnoreCase(valigment)){
			return 3;
		}else{
			throw new ZerodocException("不支持的垂直对齐方式：" + valigment);
		}
	}

	static final public int wdCellAlignVerticalTop = 0;
	static final public int wdCellAlignVerticalCenter = 1;
	static final public int wdCellAlignVerticalBottom = 3;

}
