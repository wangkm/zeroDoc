package com.zeroapp.zerodoc.msword.util;

public class LineType
{
	public static int getStyleID(String style) throws Exception
	{
		if (style.equalsIgnoreCase("wdLineStyleNone"))
			return 0;
		else if (style.equalsIgnoreCase("wdLineStyleSingle"))
			return 1;
		else if (style.equalsIgnoreCase("wdLineStyleDot"))
			return 2;
		else if (style.equalsIgnoreCase("wdLineStyleDashSmallGap"))
			return 3;
		else if (style.equalsIgnoreCase("wdLineStyleDashLargeGap"))
			return 4;
		else if (style.equalsIgnoreCase("wdLineStyleDashDot"))
			return 5;
		else if (style.equalsIgnoreCase("wdLineStyleDashDotDot"))
			return 6;
		else if (style.equalsIgnoreCase("wdLineStyleDouble"))
			return 7;
		else if (style.equalsIgnoreCase("wdLineStyleTriple"))
			return 8;
		else if (style.equalsIgnoreCase("wdLineStyleThinThickSmallGap"))
			return 9;
		else if (style.equalsIgnoreCase("wdLineStyleThickThinSmallGap"))
			return 10;
		else if (style.equalsIgnoreCase("wdLineStyleThinThickThinSmallGap"))
			return 11;
		else if (style.equalsIgnoreCase("wdLineStyleThinThickMedGap"))
			return 12;
		else if (style.equalsIgnoreCase("wdLineStyleThickThinMedGap"))
			return 13;
		else if (style.equalsIgnoreCase("wdLineStyleThinThickThinMedGap"))
			return 14;
		else if (style.equalsIgnoreCase("wdLineStyleThinThickLargeGap"))
			return 15;
		else if (style.equalsIgnoreCase("wdLineStyleThickThinLargeGap"))
			return 16;
		else if (style.equalsIgnoreCase("wdLineStyleThinThickThinLargeGap"))
			return 17;
		else if (style.equalsIgnoreCase("wdLineStyleSingleWavy"))
			return 18;
		else if (style.equalsIgnoreCase("wdLineStyleDoubleWavy"))
			return 19;
		else if (style.equalsIgnoreCase("wdLineStyleDashDotStroked"))
			return 20;
		else if (style.equalsIgnoreCase("wdLineStyleEmboss3D"))
			return 21;
		else if (style.equalsIgnoreCase("wdLineStyleEngrave3D"))
			return 22;
		else if (style.equalsIgnoreCase("wdLineStyleOutset"))
			return 23;
		else if (style.equalsIgnoreCase("wdLineStyleInset"))
			return 24;
		else
			throw new Exception("不支持的线型：" + style + "！");
	}
		
}
