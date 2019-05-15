package com.zeroapp.zerodoc.ZInterface;

public interface ZImageProvider extends ZProvider
{

	/**
	 * 返回图片的路径 返回值为Object，
	 * 如果是String，则为图片文件路径； 
	 * 如果是byte[]，则为图片文件数据
	 * 
	 * @return
	 */
	public Object get_image();

}
