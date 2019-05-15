package com.zeroapp.zerodoc.msword.util;

public class PageOrientation
{

	static public int getOrientID(String orient) throws Exception
	{
		if (orient.equalsIgnoreCase("wdOrientLandscape"))
		{
			return 1;
		}
		else if (orient.equalsIgnoreCase("wdOrientPortrait"))
		{
			return 0;
		}
		else
		{
			throw new Exception("非法的页面方向: " + orient);
		}
	}
}
