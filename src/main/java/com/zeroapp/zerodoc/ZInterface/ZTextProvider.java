package com.zeroapp.zerodoc.ZInterface;

import java.util.List;

import com.zeroapp.zerodoc.common.Data_unit;

public interface ZTextProvider extends ZProvider
{
	/**
	 * 返回结果字符串
	 * 
	 * @return
	 */
	public List<Data_unit> get_text();

}
