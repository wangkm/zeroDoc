package com.zeroapp.zerodoc.msword.util;

import org.dom4j.Element;

public class FontStyle {
	private String font_name = null;
	private Float font_size = null;
	private Long color = Color.RGB(0, 0, 0);
	private Boolean bold = false;
	private Boolean italic = false;
	private Boolean underline = false;
	private Float font_spacing = null;

	public FontStyle(Element style_info) throws Exception {
		getStyleFromXML(style_info);
	}

	/**
	 * 从Xml中提取style信息
	 * 
	 * @param style_info
	 * @throws Exception
	 */
	public void getStyleFromXML(Element style_info) throws Exception {
		if (style_info == null)
			return;

		String font_name = style_info.attributeValue("font-name");
		String font_size = style_info.attributeValue("font-size");
		String font_color = style_info.attributeValue("font-color");
		String font_weight = style_info.attributeValue("font-weight");
		String is_italic = style_info.attributeValue("italic");
		String is_underline = style_info.attributeValue("underline");
		String font_spacing = style_info.attributeValue("font-spacing");
		

		// 字体名称
		if (font_name != null)
			setFont_name(font_name);

		// 字体大小
		if (font_size != null)
			setFont_size(Float.parseFloat(font_size));

		// 字体颜色，缺省为黑色
		if (font_color != null) {
			String[] rgb = font_color.split(",");
			if(rgb.length != 3){
				System.out.println("字体颜色设置有误，格式应为用逗号分开的rgb颜色值，如：0,0,255");
			}else{
				try{
					setColor(Integer.parseInt(rgb[0].trim()), Integer.parseInt(rgb[1].trim()),
							Integer.parseInt(rgb[2].trim()));
				}catch(Exception e){
					System.out.println("设定字体颜色失败：" + e.getMessage());
				}
			}
		}else{
			setColor(0, 0, 0);
		}

		// 是否加粗，缺省为否
		setBold(false);
		if ("bold".equalsIgnoreCase(font_weight)) {
			setBold(true);
		}

		// 是否斜体，缺省为否
		setItalic(false);
		if("true".equalsIgnoreCase(is_italic)){
			setItalic(true);
		}

		// 是否有下划线，缺省为否
		setUnderline(false);
		if ("true".equalsIgnoreCase(is_underline)) {
			setUnderline(true);
		}
		
		// 字符间距
		if (font_spacing != null)
			this.setFont_spacing(Float.parseFloat(font_spacing));

	}

	public Long getColor() {
		return color;
	}

	public void setColor(int red, int green, int blue) {
		this.color = Color.RGB(red, green, blue);
	}

	public String getFont_name() {
		return font_name;
	}

	public void setFont_name(String font_name) {
		this.font_name = font_name;
	}

	public Float getFont_size() {
		return font_size;
	}

	public void setFont_size(Float font_size) {
		this.font_size = font_size;
	}

	public Boolean isBold() {
		return bold;
	}

	public void setBold(boolean bold) {
		this.bold = bold;
	}

	public Boolean isItalic() {
		return italic;
	}

	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	public Boolean isUnderline() {
		return underline;
	}

	public void setUnderline(boolean underline) {
		this.underline = underline;
	}

	public Float getFont_spacing() {
		return font_spacing;
	}

	public void setFont_spacing(Float font_spacing) {
		this.font_spacing = font_spacing;
	}
}
