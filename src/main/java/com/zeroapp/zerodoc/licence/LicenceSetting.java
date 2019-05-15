package com.zeroapp.zerodoc.licence;

public class LicenceSetting
{
	/*
	 * 版本列表 TRIAL 试用版 无时间和ip限制，最多支持20个节点 STANDARD 标准版 无时间和ip限制，最多支持1000个节点
	 */
	public enum LicTypeEnum
	{
		TRIAL, STANDARD, ENTERPRISE
	}

	/**
	 * 设定不同许可证可以支持的最大节点数目
	 * @author wkm
	 *
	 */
	public class MaxElementCount
	{
		public static final int TRIAL_MAX = 100;
		public static final int STAND_MAX = 10000;
		public static final int ENTERPRISE_MAX = 1000000;	//Integer.MAX_VALUE;
	}
}
