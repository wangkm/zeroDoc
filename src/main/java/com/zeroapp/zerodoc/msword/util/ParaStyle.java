package com.zeroapp.zerodoc.msword.util;

import org.dom4j.Element;

import com.zeroapp.zerodoc.exceptions.ZerodocException;

public class ParaStyle
{
	private Float SpaceBefore = (float)0;
	private Float SpaceAfter = (float)0;
	private Float lineSpacing = null;
	private Integer lineSpacingRule = 0;
	private Integer alignment = AlignmentType.wdAlignParagraphLeft;

	public ParaStyle(Element style_info) throws Exception{
		getStyleFromXML(style_info);
	}
	public void getStyleFromXML(Element style_info) throws Exception
	{
		if(style_info == null){
			return;
		}
		
		// 对齐方式，缺省为左对其
		String alignment = style_info.attributeValue("alignment");
		if (alignment != null){
			this.setAlignment(AlignmentType.getAlignmentID(alignment));
		}else{
			this.setAlignment(AlignmentType.wdAlignParagraphLeft);
		}

		// 段前间距，单位为磅，缺省为0
		String SpaceBefore_s = style_info.attributeValue("spaceBefore");
		if (SpaceBefore_s != null){
			this.setSpaceBefore(Float.parseFloat(SpaceBefore_s));
		}else{
			this.setSpaceBefore((float)0);
		}

		// 段后间距，单位为磅，缺省为0
		String SpaceAfter_s = style_info.attributeValue("spaceAfter");
		if (SpaceAfter_s != null){
			this.setSpaceAfter(Float.parseFloat(SpaceAfter_s));
		}else{
			this.setSpaceAfter((float)0);
		}

		// 段落间距样式，缺省为单倍行距
		String lineSpacingRule = style_info.attributeValue("lineSpacingRule");
		if (lineSpacingRule != null){
			this.setLineSpacingRule(getLineSpacingRuleID(lineSpacingRule));
		}else{
			this.setLineSpacingRule(0);
		}

		// 段间距离，单位为磅
		String lineSpacing = style_info.attributeValue("lineSpacing");
		if (lineSpacing != null)
			this.setLineSpacing(Float.parseFloat(lineSpacing));

	}

	public static int getLineSpacingRuleID(String LineSpacingRule)
			throws Exception
	{
		if (LineSpacingRule.equalsIgnoreCase("wdLineSpaceSingle"))
			return 0;
		else if (LineSpacingRule.equalsIgnoreCase("wdLineSpace1pt5"))
			return 1;
		else if (LineSpacingRule.equalsIgnoreCase("wdLineSpaceDouble"))
			return 2;
		else if (LineSpacingRule.equalsIgnoreCase("wdLineSpaceAtLeast"))
			return 3;
		else if (LineSpacingRule.equalsIgnoreCase("wdLineSpaceExactly"))
			return 4;
		else if (LineSpacingRule.equalsIgnoreCase("wdLineSpaceMultiple"))
			return 5;
		else
			throw new ZerodocException("非法的段落间距定义: " + LineSpacingRule + "！");
	}


	public Float getLineSpacing()
	{
		return lineSpacing;
	}

	public void setLineSpacing(float lineSpacing)
	{
		this.lineSpacing = lineSpacing;
	}

	public Integer getLineSpacingRule()
	{
		return lineSpacingRule;
	}

	public void setLineSpacingRule(int lineSpacingRule)
	{
		this.lineSpacingRule = lineSpacingRule;
	}

	public Integer getAlignment()
	{
		return alignment;
	}

	public void setAlignment(int alignment)
	{
		this.alignment = alignment;
	}
	public Float getSpaceBefore() {
		return SpaceBefore;
	}
	public void setSpaceBefore(Float spaceBefore) {
		SpaceBefore = spaceBefore;
	}
	public Float getSpaceAfter() {
		return SpaceAfter;
	}
	public void setSpaceAfter(Float spaceAfter) {
		SpaceAfter = spaceAfter;
	}

}
