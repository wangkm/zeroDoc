package com.zeroapp.zerodoc.msword;

import org.dom4j.Element;

import com.jacob2.com.Dispatch;
import com.jacob2.com.Variant;
import com.zeroapp.zerodoc.msword.util.BuildInStyle;
import com.zeroapp.zerodoc.msword.util.FontStyle;
import com.zeroapp.zerodoc.msword.util.ParaStyle;
import com.zeroapp.zerodoc.util.MiscUtil;

/**
 * 文档中指定的一段连续区域
 * 
 * @author wkm
 * 
 */
public class Range {
	Dispatch _range = null;

	public Range(Variant range) {
		_range = range.getDispatch();
	}

	public Range(Dispatch range) {
		_range = range;
	}

	public void appendText(String text, Element style_info) throws Exception {

		try {
			Dispatch.call(_range, "InsertAfter", text);
			this.setStyle(style_info);
		} catch (Exception e) {
			System.out.println(text);
			throw e;
		}
	}
	
	public void setStyle(Element style_info) throws Exception{
		if(style_info == null){
			return;
		}
		
		// 设置当前区域的内建样式
		String buildInStyle = style_info.attributeValue("buildin-style");
		if(buildInStyle != null){
			int buildinStyle = BuildInStyle.getStyleID(buildInStyle);
			this.setBuildinStyle(buildinStyle);
		}

		// 设置当前区域的段落样式
		ParaStyle paraStyle = new ParaStyle(style_info);
		this.setParaStyle(paraStyle);

		// 设置当前区域的字体样式
		FontStyle fontStyle = new FontStyle(style_info);
		this.setFontStyle(fontStyle);
		
	}

	/**
	 * 设定当前选中区域的段落样式
	 * 
	 * @param paraStyle
	 * @throws Exception
	 */
	private void setParaStyle(ParaStyle paraStyle) throws Exception {
		if (paraStyle == null) {
			return;
		}

		Dispatch paragraphFormat = Dispatch.get(_range, "ParagraphFormat")
				.toDispatch();

		if (paraStyle.getAlignment() != null) {
			Dispatch.put(paragraphFormat, "Alignment", new Variant(paraStyle
					.getAlignment()));
		}
		if (paraStyle.getSpaceBefore() != null) {
			Dispatch.put(paragraphFormat, "SpaceBefore", new Variant(
					paraStyle.getSpaceBefore()));
		}
		if (paraStyle.getSpaceAfter() != null) {
			Dispatch.put(paragraphFormat, "SpaceAfter", new Variant(
					paraStyle.getSpaceAfter()));
		}
		if (paraStyle.getLineSpacingRule() != null) {
			Dispatch.put(paragraphFormat, "LineSpacingRule", new Variant(
					paraStyle.getLineSpacingRule()));
		}
		if (paraStyle.getLineSpacing() != null) {
			float spacing = paraStyle.getLineSpacing();
			if(paraStyle.getLineSpacingRule() == 5 ){
				spacing = MiscUtil.LinesToPoints(spacing);
			}
			Dispatch.put(paragraphFormat, "LineSpacing", new Variant(spacing));
		}

	}

	/**
	 * 设定字体样式
	 * 
	 * @param fontStyle
	 */
	private void setFontStyle(FontStyle fontStyle) {
		if (fontStyle == null) {
			return;
		}

		Dispatch font = Dispatch.get(_range, "Font").toDispatch();
		
		// 字体名称
		if (fontStyle.getFont_name() != null)
			Dispatch.put(font, "Name", new Variant(fontStyle.getFont_name()));
		
		// 是否粗体
		if (fontStyle.isBold() != null)
			Dispatch.put(font, "Bold", new Variant(fontStyle.isBold()));
		
		// 是否斜体
		if (fontStyle.isItalic() != null)
			Dispatch.put(font, "Italic", new Variant(fontStyle.isItalic()));
		
		// 是否有下划线
		if (fontStyle.isUnderline() != null)
			Dispatch.put(font, "Underline",
					new Variant(fontStyle.isUnderline()));
		
		// 字体颜色
		if (fontStyle.getColor() != null)
			Dispatch.put(font, "Color",
					new Variant((long) fontStyle.getColor()));
		
		// 字体大小
		if (fontStyle.getFont_size() != null)
			Dispatch.put(font, "Size", new Variant(fontStyle.getFont_size()));
		
		// 字符间距
		if (fontStyle.getFont_spacing() != null)
			Dispatch.put(font, "Spacing", new Variant(fontStyle.getFont_spacing()));

	}

	/**
	 * 设定选中部分的样式，如标题1，标题2等等
	 * 
	 * @param n
	 * @throws Exception
	 */
	private void setBuildinStyle(int buildinStyle) throws Exception {
		Dispatch.put(_range, "Style", new Variant(buildinStyle));
	}

	/**
	 * 设定为缺省的style
	 * 
	 * @throws Exception
	 */
	public void setDefaultStyle() throws Exception {
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		Dispatch paragraphFormat = Dispatch.get(selection, "ParagraphFormat")
				.toDispatch();
		Dispatch.put(paragraphFormat, "LineSpacingRule", new Variant(ParaStyle
				.getLineSpacingRuleID("wdLineSpaceSingle")));

		this.setBuildinStyle(BuildInStyle
				.getStyleID("wdStyleDefaultParagraphFont"));

	}

}
