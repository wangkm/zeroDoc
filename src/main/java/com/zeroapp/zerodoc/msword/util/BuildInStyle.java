package com.zeroapp.zerodoc.msword.util;

public class BuildInStyle
{
	public static int getStyleID(String style) throws Exception
	{
		if (style.equalsIgnoreCase("wdStyleHeading1"))
			return -2;
		else if (style.equalsIgnoreCase("wdStyleHeading2"))
			return -3;
		else if (style.equalsIgnoreCase("wdStyleHeading3"))
			return -4;
		else if (style.equalsIgnoreCase("wdStyleHeading4"))
			return -5;
		else if (style.equalsIgnoreCase("wdStyleHeading5"))
			return -6;
		else if (style.equalsIgnoreCase("wdStyleHeading6"))
			return -7;
		else if (style.equalsIgnoreCase("wdStyleDefaultParagraphFont"))
			return -66;
		else
			throw new Exception("不支持的内建样式：" + style + "！");
	}

	static int wdStyleHeading1 = -2;
	static int wdStyleHeading2 = -3;
	static int wdStyleHeading3 = -4;
	static int wdStyleHeading4 = -5;
	static int wdStyleHeading5 = -6;
	static int wdStyleHeading6 = -7;
	static int wdStyleDefaultParagraphFont = -66;
}
