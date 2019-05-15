package com.zeroapp.zerodoc.util;

public class MiscUtil {
	/**
	 * 将度量单位由厘米转换为磅数（1 厘米=28.35 磅）。 所返回的转换结果为 double 类型。
	 */
	static public double centimetersToPoints(double centimeters) {
		return centimeters * 28.35;
	}

	/**
	 * 将度量单位从行转换为磅（1 行 = 12 磅）。
	 * @param lines
	 * @return
	 */
	static public float LinesToPoints(float lines){
		return lines * 12;
	}
}
