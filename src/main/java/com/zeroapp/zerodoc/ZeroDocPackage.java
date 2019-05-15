package com.zeroapp.zerodoc;

public class ZeroDocPackage
{
	// 设定产品名称
	private static String _product_name = "ZeroDoc";

	// 设定版本号
	private static String _version = "1.4";

	/**
	 * 获取产品的名称
	 * @return
	 */
	public static String get_product_name()
	{
		return _product_name;
	}

	/**
	 * 获取产品当前版本号
	 * @return
	 */
	public static String get_version()
	{
		return _version;
	}

}
