package com.zeroapp.zerodoc.ZInterface;

import java.util.List;

import com.zeroapp.zerodoc.common.Parameter;

public interface ZProvider
{
	/**
	 * 回调函数
	 * 获取调用时候的特殊设定
	 * @return
	 */
	public String set_options(String options);
	
	/**
	 * 回调函数，获取参数列表，
	 * 每个参数包括参数名、参数缺省值、参数类型，参数说明
	 * 
	 * @return
	 */
	public List<Parameter> get_para_info();

	/**
	 * 回调函数
	 * 设定参数
	 */
	public void set_parameters(List<Parameter> para_list);

}
