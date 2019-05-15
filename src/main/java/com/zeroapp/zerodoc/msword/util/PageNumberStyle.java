package com.zeroapp.zerodoc.msword.util;

import com.zeroapp.zerodoc.exceptions.ZerodocException;

public class PageNumberStyle {
	// 页码的形式
	static final public int wdPageNumberStyleArabic = 0;
	static final public int wdPageNumberStyleUppercaseRoman = 1;
	static final public int wdPageNumberStyleLowercaseRoman = 2;
	static final public int wdPageNumberStyleUppercaseLetter = 3;
	static final public int wdPageNumberStyleLowercaseLetter = 4;
	static final public int wdPageNumberStyleSimpChinNum1 = 37;
	static final public int wdPageNumberStyleSimpChinNum2 = 38;
	static final public int wdPageNumberStyleTradChinNum1 = 33;
	static final public int wdPageNumberStyleTradChinNum2 = 34;

	
	public static int getPageNumberStyleID(String style) throws ZerodocException{
		if("wdPageNumberStyleArabic".equalsIgnoreCase(style)){
			return wdPageNumberStyleArabic;
		}else if("wdPageNumberStyleUppercaseRoman".equalsIgnoreCase(style)){
			return wdPageNumberStyleUppercaseRoman;
		}else if("wdPageNumberStyleLowercaseRoman".equalsIgnoreCase(style)){
			return wdPageNumberStyleLowercaseRoman;
		}else if("wdPageNumberStyleUppercaseLetter".equalsIgnoreCase(style)){
			return wdPageNumberStyleUppercaseLetter;
		}else if("wdPageNumberStyleLowercaseLetter".equalsIgnoreCase(style)){
			return wdPageNumberStyleLowercaseLetter;
		}else if("wdPageNumberStyleSimpChinNum1".equalsIgnoreCase(style)){
			return wdPageNumberStyleSimpChinNum1;
		}else if("wdPageNumberStyleSimpChinNum2".equalsIgnoreCase(style)){
			return wdPageNumberStyleSimpChinNum2;
		}else if("wdPageNumberStyleTradChinNum1".equalsIgnoreCase(style)){
			return wdPageNumberStyleTradChinNum1;
		}else if("wdPageNumberStyleTradChinNum2".equalsIgnoreCase(style)){
			return wdPageNumberStyleTradChinNum2;
		}else{
			throw new ZerodocException("不正确的页码格式：" + style);
		}
	}
}
