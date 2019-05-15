package com.zeroapp.zerodoc.msword;

import org.dom4j.Element;

import com.jacob2.com.Dispatch;
import com.jacob2.com.Variant;
import com.zeroapp.zerodoc.msword.util.ParaStyle;

public class Image {
	Dispatch _image = null;

	/**
	 * 构造函数
	 * 
	 * @param image
	 */
	public Image(Variant image) {
		_image = image.getDispatch();
	}

	/**
	 * 设定图片高度
	 * 
	 * @param height
	 */
	public void setHeight(int height) {
		Dispatch.put(_image, "Height", new Variant(height));
	}

	/**
	 * 设定图片宽度
	 * 
	 * @param Width
	 */
	public void setWidth(int Width) {
		Dispatch.put(_image, "Width", new Variant(Width));
	}

	/**
	 * 设定图片位置
	 */
	public void setAlignment(Integer alignment) {
		if (alignment != null) {
			// 得到selection对象
			Dispatch.call(_image, "Select");
			Dispatch selection = Dispatch.get(WordFactory.MsWordApp,
					"Selection").toDispatch();
			Dispatch paragraphFormat = Dispatch.get(selection,
					"ParagraphFormat").toDispatch();
			Dispatch.put(paragraphFormat, "Alignment", new Variant(alignment));
		}
	}

	/**
	 * 设定图片的样式
	 * 
	 * @param style_info
	 * @throws Exception 
	 */
	public void setStyle(Element style_info) throws Exception {
		if (style_info == null) {
			return;
		}
		
		ParaStyle paraStyle = new ParaStyle(style_info);
		this.setAlignment(paraStyle.getAlignment());

		String image_height = style_info.attributeValue("height");
		if (image_height != null) {
			this.setHeight(Integer.parseInt(image_height));
		}

		String image_width = style_info.attributeValue("width");
		if (image_width != null) {
			this.setWidth(Integer.parseInt(image_width));
		}
	}

}
