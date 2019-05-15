package com.zeroapp.zerodoc.msword.util;

import com.zeroapp.zerodoc.exceptions.ZerodocException;

public class AlignmentType
{

	public static int getAlignmentID(String aligment) throws Exception
	{
		if (aligment.equalsIgnoreCase("wdAlignParagraphLeft"))
			return 0;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphCenter"))
			return 1;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphDistribute"))
			return 4;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphJustify"))
			return 3;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphJustifyHi"))
			return 7;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphJustifyLow"))
			return 8;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphJustifyMed"))
			return 5;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphRight"))
			return 2;
		else if (aligment.equalsIgnoreCase("wdAlignParagraphThaiJustify"))
			return 9;
		else
			throw new ZerodocException("不支持的对齐方式：" + aligment);
	}

	public static final int wdAlignParagraphCenter = 1;
	public static final int wdAlignParagraphDistribute = 4;
	public static final int wdAlignParagraphJustify = 3;
	public static final int wdAlignParagraphJustifyHi = 7;
	public static final int wdAlignParagraphJustifyLow = 8;
	public static final int wdAlignParagraphJustifyMed = 5;
	public static final int wdAlignParagraphLeft = 0;
	public static final int wdAlignParagraphRight = 2;
	public static final int wdAlignParagraphThaiJustify = 9;

}
