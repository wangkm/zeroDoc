package com.zeroapp.zerodoc.common;

public class Parameter
{
	private String name;	// 参数名
	private String value;	// 参数值
	private String type;	// 参数类型
	private String remark;	// 参数说明

	public String getType()
	{
		return type;
	}

	public void setType(String type)
	{
		this.type = type;
	}

	/**
	 * 构造函数
	 * 
	 * @param name
	 * @param value
	 * @param remark
	 */
	public Parameter(String name, String value, String remark)
	{
		setName(name);
		setValue(value);
		setRemark(remark);
	}

	public String getName()
	{
		return name;
	}

	public void setName(String name)
	{
		this.name = name;
	}

	public String getValue()
	{
		return value;
	}

	public void setValue(String value)
	{
		this.value = value;
	}

	public String getRemark()
	{
		return remark;
	}

	public void setRemark(String remark)
	{
		this.remark = remark;
	}

}
