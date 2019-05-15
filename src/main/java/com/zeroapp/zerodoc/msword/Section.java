package com.zeroapp.zerodoc.msword;

import org.dom4j.Element;

import com.jacob2.activeX.ActiveXComponent;
import com.jacob2.com.Dispatch;
import com.jacob2.com.Variant;
import com.zeroapp.zerodoc.exceptions.ZerodocException;
import com.zeroapp.zerodoc.msword.util.BorderStyle;
import com.zeroapp.zerodoc.msword.util.ConstantUtil;
import com.zeroapp.zerodoc.msword.util.PageOrientation;
import com.zeroapp.zerodoc.msword.util.PageNumberStyle;
import com.zeroapp.zerodoc.util.MiscUtil;

public class Section {
	Dispatch _section = null;

	/**
	 * 构造函数
	 * 
	 * @param table
	 */
	public Section(Variant section) {
		_section = section.getDispatch();
	}

	/**
	 * 设定页眉是否关联到前一节
	 * 
	 * @param LinkToPrevious
	 */
	public void setHeaderLinkToPrevious(boolean LinkToPrevious) {
		Dispatch headers = ActiveXComponent.call(_section, "Headers",
				new Variant(ConstantUtil.wdHeaderFooterPrimary)).toDispatch();
		Dispatch.put(headers, "LinkToPrevious", LinkToPrevious);
	}

	/**
	 * 设定页脚是否关联到前一节
	 * 
	 * @param LinkToPrevious
	 */
	public void setFootLinkToPrevious(boolean LinkToPrevious) {
		Dispatch footer = ActiveXComponent.call(_section, "Footers",
				new Variant(ConstantUtil.wdHeaderFooterPrimary)).toDispatch();
		Dispatch.put(footer, "LinkToPrevious", LinkToPrevious);
	}

	/**
	 * 清除页眉内容
	 */
	public void clearHeader() {
		Dispatch headers = ActiveXComponent.call(_section, "Headers",
				new Variant(ConstantUtil.wdHeaderFooterPrimary)).toDispatch();
		Dispatch range = Dispatch.call(headers, "Range").toDispatch();
		Dispatch.call(range, "Delete");

	}

	/**
	 * 清除页脚内容
	 */
	public void clearFooter() {
		Dispatch footer = ActiveXComponent.call(_section, "Footers",
				new Variant(ConstantUtil.wdHeaderFooterPrimary)).toDispatch();
		Dispatch range = Dispatch.call(footer, "Range").toDispatch();
		Dispatch.call(range, "Delete");

	}

	/**
	 * 设定本段落中的页码格式
	 * 
	 * @param restart
	 * @param startingNumber
	 * @param numberStyle
	 * @param showFirstPageNumber
	 * @throws Exception
	 */
	public void setPageNumberProperity(String restart,
			String startingNumber, String numberStyle,
			String showFirstPageNumber) throws Exception {
		Dispatch headers = Dispatch.call(_section, "Headers",
				new Variant(ConstantUtil.wdHeaderFooterPrimary)).toDispatch();
		Dispatch pageNumbers = Dispatch.call(headers, "PageNumbers")
				.toDispatch();

		try {
			if (restart != null) {
				Dispatch.put(pageNumbers, "RestartNumberingAtSection",
						new Variant(Boolean.parseBoolean(restart)));
			}
			if (startingNumber != null) {
				Dispatch.put(pageNumbers, "StartingNumber", new Variant(Integer
						.parseInt(startingNumber)));
			}
			if (numberStyle != null) {
				Dispatch.put(pageNumbers, "NumberStyle", new Variant(PageNumberStyle.getPageNumberStyleID(numberStyle)));
			}
			if (showFirstPageNumber != null) {
				Dispatch.put(pageNumbers, "ShowFirstPageNumber", new Variant(
						Boolean.parseBoolean(showFirstPageNumber)));
			}
		} catch (Exception e) {
			throw new Exception("设定页眉上的页码格式出错：" + e.getMessage());
		}
	}

	/**
	 * 设定页面的方向 只对当前的节有效，不影响其他节的设定
	 * 
	 * @param orientType
	 * @throws Exception
	 */
	public void setPageOrientation(int orientType) throws Exception {
		Dispatch pagesetup = Dispatch.call(_section, "PageSetup").toDispatch();

		Dispatch.put(pagesetup, "Orientation", new Variant(orientType));
	}

	/**
	 * 设定当前节的分栏数目
	 * 
	 * @param count
	 * @throws ZerodocException
	 */
	public void setTextColumns(int count) throws ZerodocException {
		if (count < 1 || count > 5) {
			throw new ZerodocException("分栏数目必须在1和5之间");
		}
		Dispatch pagesetup = Dispatch.call(_section, "PageSetup").toDispatch();
		Dispatch textColumns = Dispatch.call(pagesetup, "TextColumns")
				.toDispatch();
		Dispatch.call(textColumns, "SetCount", new Variant(count));
	}

	/**
	 * 去掉页眉上的横线
	 */
	public void removeHeaderLine() {
		Dispatch header = Dispatch.call(_section, "Headers",
				new Variant(ConstantUtil.wdHeaderFooterPrimary)).toDispatch();
		Dispatch range = Dispatch.get(header, "Range").toDispatch();
		Dispatch paragraphFormat = Dispatch.get(range, "ParagraphFormat")
				.toDispatch();
		Dispatch borders = Dispatch.get(paragraphFormat, "Borders")
				.toDispatch();
		Dispatch border = Dispatch.call(borders, "Item",
				new Variant(BorderStyle.wdBorderBottom)).toDispatch();
		Dispatch.put(border, "LineStyle", new Variant(0));

	}
	
	/**
	 * 设定本节的属性
	 * @param page_setup
	 * @throws ZerodocException
	 */
	public void setPageProperties(Element section) throws ZerodocException {
		// 设定页面的高度和宽度
		setPageWidthAndHeight(section);

		// 设定文档的四个边距
		setPageMargin(section);

		// 设定页眉页脚与页面上下边之间的距离
		setFHDistance(section);
		
		// 设定本节的页面方向
		String pageOrientation = section.attributeValue("pageOrientation");
		// if page orientation setting is valid, set it
		if (pageOrientation != null) {
			try {
				this.setPageOrientation(PageOrientation
						.getOrientID(pageOrientation));
			} catch (Exception e) {
				System.out.println("设定页面方向失败：" + e.getMessage());
			}
		}

		// 设定分栏
		String columnCount = section.attributeValue("column_count");
		// if page orientation setting is valid, set it
		if (columnCount != null) {
			try {
				this.setTextColumns(Integer.parseInt(columnCount));
			} catch (Exception e) {
				System.out.println("设定分栏失败：" + e.getMessage());
			}
		}

		// 设定本节的页码格式
		String pagenum_restart = section
				.attributeValue("pagenum_restart");
		String pagenum_startingnum = section
				.attributeValue("pagenum_startingnum");
		String pagenum_numstyle = section
				.attributeValue("pagenum_numstyle");
		String pagenum_showfirst = section
				.attributeValue("pagenum_showfirst");
		
		try {
			this.setPageNumberProperity(pagenum_restart,
					pagenum_startingnum, pagenum_numstyle,
					pagenum_showfirst);
		} catch (Exception e) {
			System.out.println("设定页码格式失败：" + e.getMessage());
		}


	}

	/**
	 * 设定页面的宽度和高度 单位为厘米
	 * 
	 * @param width
	 * @param height
	 * @throws ZerodocException
	 */
	public void setPageWidthAndHeight(final Double pageWidth,
			final Double pageHeight) throws ZerodocException {
		Dispatch pagesetup = ActiveXComponent.call(_section, "PageSetup")
				.toDispatch();
		
		if (pageWidth != null) {
			if (MiscUtil.centimetersToPoints(pageWidth) < 7.2
					|| MiscUtil.centimetersToPoints(pageWidth) > 1584) {
				throw new ZerodocException("页面宽度设置非法");
			}
			Dispatch.put(pagesetup, "pageWidth", new Variant(
					MiscUtil.centimetersToPoints(pageWidth)));
		}
		if (pageHeight != null) {
			if (MiscUtil.centimetersToPoints(pageHeight) < 7.2
					|| MiscUtil.centimetersToPoints(pageHeight) > 1584) {
				throw new ZerodocException("页面高度设置非法");
			}
			Dispatch.put(pagesetup, "pageHeight", new Variant(
					MiscUtil.centimetersToPoints(pageHeight)));
		}

	}
	/**
	 * 设定页眉与页面顶边之间的距离 设定页脚与页底边之间的距离 以下两个参数分别对应两个距离 如果某个参数为NULL，则采用缺省边距
	 * 
	 * @param HeaderDistance
	 * @param FooterDistance
	 */
	public void setFHDistance(final Double headerDistance,
			final Double footerDistance) {
		double realHeaderDistance = (headerDistance != null ? headerDistance
				: ConstantUtil.default_headerDistance);
		double realFooterDistance = (headerDistance != null ? footerDistance
				: ConstantUtil.default_footerDistance);

		Dispatch pagesetup = ActiveXComponent.call(_section, "PageSetup")
				.toDispatch();

		Dispatch.put(pagesetup, "HeaderDistance", new Variant(
				MiscUtil.centimetersToPoints(realHeaderDistance)));
		Dispatch.put(pagesetup, "FooterDistance", new Variant(
				MiscUtil.centimetersToPoints(realFooterDistance)));

	}
	/**
	 * 设定页面的边距 以下四个参数分别对应四个边距 如果某个参数为NULL，则采用缺省边距
	 * 
	 * @param topMargin
	 * @param bottomMargin
	 * @param leftMargin
	 * @param rightMargin
	 */
	public void setPageMargin(final Double topMargin,
			final Double bottomMargin, final Double leftMargin,
			final Double rightMargin) {
		//
		double realTopMargin = (topMargin != null ? topMargin
				: ConstantUtil.default_topMargin);
		double realBottomMargin = (bottomMargin != null ? bottomMargin
				: ConstantUtil.default_buttomMargin);
		double realLeftMargin = (leftMargin != null ? leftMargin
				: ConstantUtil.default_leftMargin);
		double realRightMargin = (rightMargin != null ? rightMargin
				: ConstantUtil.default_rightMargin);

		Dispatch pagesetup = ActiveXComponent.call(_section, "PageSetup")
				.toDispatch();

		Dispatch.put(pagesetup, "TopMargin", new Variant(
				MiscUtil.centimetersToPoints(realTopMargin)));
		Dispatch.put(pagesetup, "BottomMargin", new Variant(
				MiscUtil.centimetersToPoints(realBottomMargin)));
		Dispatch.put(pagesetup, "LeftMargin", new Variant(
				MiscUtil.centimetersToPoints(realLeftMargin)));
		Dispatch.put(pagesetup, "RightMargin", new Variant(
				MiscUtil.centimetersToPoints(realRightMargin)));

	}
	/**
	 * 设定页面的四个页边距
	 * 
	 * @param pageSetup
	 */
	private void setPageMargin(Element page_setup) {
		if (page_setup == null) {
			return;
		}
		// 获取上面的页边距
		String strTopMargin = page_setup.attributeValue("TopMargin");
		Double topMargin = null;
		try {
			topMargin = Double.parseDouble(strTopMargin);
			if (topMargin < 0) {
				topMargin = null;
			}
		} catch (Exception e) {
			topMargin = null;
		}

		// 获取下面的页边距
		String strBottomMargin = page_setup.attributeValue("BottomMargin");
		Double bottomMargin = null;
		try {
			bottomMargin = Double.parseDouble(strBottomMargin);
			if (bottomMargin < 0) {
				bottomMargin = null;
			}
		} catch (Exception e) {
			bottomMargin = null;
		}

		// 获取左面的页边距
		String strLeftMargin = page_setup.attributeValue("LeftMargin");
		Double leftMargin = null;
		try {
			leftMargin = Double.parseDouble(strLeftMargin);
			if (leftMargin < 0) {
				leftMargin = null;
			}
		} catch (Exception e) {
			leftMargin = null;
		}
		// 获取右面的页边距
		String strRightMargin = page_setup.attributeValue("RightMargin");
		Double rightMargin = null;
		try {
			rightMargin = Double.parseDouble(strRightMargin);
			if (rightMargin < 0) {
				rightMargin = null;
			}
		} catch (Exception e) {
			rightMargin = null;
		}

		// 设定页边距
		this.setPageMargin(topMargin, bottomMargin, leftMargin,
				rightMargin);
	}

	/**
	 * 设定页眉与页面顶边之间的距离 设定页脚与页底边之间的距离
	 * 
	 * @param page_setup
	 */
	private void setFHDistance(Element page_setup) {
		if (page_setup == null) {
			return;
		}
		// 获取页眉与页面顶边之间的距离
		String strHeaderDistance = page_setup.attributeValue("HeaderDistance");
		Double headerDistance = null;
		try {
			headerDistance = Double.parseDouble(strHeaderDistance);
			if (headerDistance < 0) {
				headerDistance = null;
			}
		} catch (Exception e) {
			headerDistance = null;
		}

		// 获取页脚与页底边之间的距离
		String strFooterDistance = page_setup.attributeValue("FooterDistance");
		Double footerDistance = null;
		try {
			footerDistance = Double.parseDouble(strFooterDistance);
			if (footerDistance < 0) {
				footerDistance = null;
			}
		} catch (Exception e) {
			footerDistance = null;
		}

		// 调用设定函数
		this.setFHDistance(headerDistance, footerDistance);
	}

	/**
	 * 设定页面的高度和宽度
	 * 
	 * @param page_setup
	 * @throws ZerodocException
	 */
	private void setPageWidthAndHeight(Element page_setup)
			throws ZerodocException {
		if (page_setup == null) {
			return;
		}

		Double pageHeight = null;
		Double pageWidth = null;
		// 首先查看是否设定了页面样式
		String pageType = page_setup.attributeValue("PageType");
		if ("A3".equalsIgnoreCase(pageType)) {
			pageWidth = 29.7;
			pageHeight = 42.0;
		} else if ("A4".equalsIgnoreCase(pageType)) {
			pageWidth = 21.0;
			pageHeight = 29.7;
		} else if ("A5".equalsIgnoreCase(pageType)) {
			pageWidth = 14.8;
			pageHeight = 21.0;
		} else if ("A6".equalsIgnoreCase(pageType)) {
			pageWidth = 10.5;
			pageHeight = 14.8;
		}
		// 否则获取指定的高度或宽度设置
		else {
			String pageWidth_s = page_setup.attributeValue("PageWidth");
			String pageHeight_s = page_setup.attributeValue("PageHeight");
			if (pageWidth_s != null) {
				try {
					pageWidth = Double.parseDouble(pageWidth_s);
				} catch (Exception e) {
					throw new ZerodocException("解析页面宽度失败：" + e.getMessage());
				}
			}
			if (pageHeight_s != null) {
				try {
					pageHeight = Double.parseDouble(pageHeight_s);
				} catch (Exception e) {
					throw new ZerodocException("解析页面高度失败：" + e.getMessage());
				}
			}
		}

		this.setPageWidthAndHeight(pageWidth, pageHeight);
	}

}
