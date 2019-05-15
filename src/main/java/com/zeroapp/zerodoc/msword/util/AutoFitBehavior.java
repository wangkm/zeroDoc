package com.zeroapp.zerodoc.msword.util;

public class AutoFitBehavior
{
	public static int getAutoFitID(String behavior) throws Exception
	{
		if (behavior.equalsIgnoreCase("wdAutoFitWindow"))
		{
			return 2;
		}
		else if (behavior.equalsIgnoreCase("wdAutoFitContent"))
		{
			return 1;
		}
		else if (behavior.equalsIgnoreCase("wdAutoFitFixed"))
		{
			return 0;
		}
		else
		{
			throw new Exception("不支持的自动对齐方式: " + behavior);
		}
	}
}
