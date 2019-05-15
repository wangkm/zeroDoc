package com.zeroapp.zerodoc.msword.util;

import org.dom4j.Element;

public class BorderStyle
{
	// 各个边框的值
	static final public int wdBorderTop = -1;
	static final public int wdBorderLeft = -2;
	static final public int wdBorderBottom = -3;
	static final public int wdBorderRight = -4;
	static final public int wdBorderHorizontal = -5;
	static final public int wdBorderVertical = -6;
	static final public int wdBorderDiagonalDown = -7;
	static final public int wdBorderDiagonalUp = -8;
	
	private Integer borderTopStyle = null;
	private Integer borderBottompStyle = null;
	private Integer borderLeftStyle = null;
	private Integer borderRightStyle = null;
	private Integer borderVerticalStyle = null;
	private Integer borderHorizontalStyle = null;
	
	public BorderStyle(Element style_info) throws Exception{
		getStyleFromXML(style_info);
	}
	
	public BorderStyle()
	{
		// 将缺省的线型设定为single
		setBorderTopStyle(1);
		setBorderBottompStyle(1);
		setBorderLeftStyle(1);
		setBorderRightStyle(1);
		setBorderHorizontalStyle(1);
		setBorderVerticalStyle(1);
		
	}

	public void getStyleFromXML(Element style_info) throws Exception{
		if(style_info == null){
			return;
		}
		
		String style;
		
		style = style_info.attributeValue("border-top");
		if(style != null){
			setBorderTopStyle(LineType.getStyleID(style));
		}
		
		style = style_info.attributeValue("border-bottom");
		if(style != null){
			setBorderBottompStyle(LineType.getStyleID(style));
		}
		
		style = style_info.attributeValue("border-left");
		if(style != null){
			setBorderLeftStyle(LineType.getStyleID(style));
		}

		style = style_info.attributeValue("border-right");
		if(style != null){
			setBorderRightStyle(LineType.getStyleID(style));
		}

		style = style_info.attributeValue("border-horizontal");
		if(style != null){
			setBorderHorizontalStyle(LineType.getStyleID(style));
		}

		style = style_info.attributeValue("border-vertical");
		if(style != null){
			setBorderVerticalStyle(LineType.getStyleID(style));
		}
	}

	public Integer getBorderTopStyle() {
		return borderTopStyle;
	}

	public void setBorderTopStyle(Integer borderTopStyle) {
		this.borderTopStyle = borderTopStyle;
	}

	public Integer getBorderBottompStyle() {
		return borderBottompStyle;
	}

	public void setBorderBottompStyle(Integer borderBottompStyle) {
		this.borderBottompStyle = borderBottompStyle;
	}

	public Integer getBorderLeftStyle() {
		return borderLeftStyle;
	}

	public void setBorderLeftStyle(Integer borderLeftStyle) {
		this.borderLeftStyle = borderLeftStyle;
	}

	public Integer getBorderRightStyle() {
		return borderRightStyle;
	}

	public void setBorderRightStyle(Integer borderRightStyle) {
		this.borderRightStyle = borderRightStyle;
	}

	public Integer getBorderVerticalStyle() {
		return borderVerticalStyle;
	}

	public void setBorderVerticalStyle(Integer borderVerticalStyle) {
		this.borderVerticalStyle = borderVerticalStyle;
	}

	public Integer getBorderHorizontalStyle() {
		return borderHorizontalStyle;
	}

	public void setBorderHorizontalStyle(Integer borderHorizontalStyle) {
		this.borderHorizontalStyle = borderHorizontalStyle;
	}

}
