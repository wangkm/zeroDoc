package com.zeroapp.zerodoc.common;

import org.dom4j.Element;

/**
 * 最小的数据单元，通过type来指定，目前支持图片和文本。
 * 
 * @author wkm
 * 
 */
public class Data_unit
{
	private String _type = null;
	private Object _data = null;
	private Element _style_info; // 样式，xml格式

	public Data_unit(String type, Object data, Element style_info)
	{
		_type = type;
		_data = data;
		_style_info = style_info;
	}

	public String get_type()
	{
		return _type;
	}

	public void set_type(String _type)
	{
		this._type = _type;
	}

	public Object get_data()
	{
		return _data;
	}

	public void set_data(Object _data)
	{
		this._data = _data;
	}

	public Element get_style_info()
	{
		return _style_info;
	}

	public void set_style_info(Element _style_info)
	{
		this._style_info = _style_info;
	}

}
