package com.zeroapp.zerodoc.msword.util;

public class Color
{

	public static long RGB(long red, long green, long blue)
	{
		long b = blue << 16;
		long g = green << 8;
		long r = red;
		return b + g + r;

	}
}
