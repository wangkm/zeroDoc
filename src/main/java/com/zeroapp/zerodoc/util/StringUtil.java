package com.zeroapp.zerodoc.util;


import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * 字符串操作相关
 * 
 */
public class StringUtil {
	
	//js中的NULL
	public static String NULL ="_NULL";
	

	/**
	 * str是否一个空字符串<br>
	 * 
	 * @param str
	 * @return
	 */
	public static boolean isEmpty(String str) {
		return str == null || str.trim().length() == 0 || str.trim().equals("　");
	}

	/**
	 * 将 true和false 翻成中文的 是和否
	 * 
	 * @param bl
	 * @return
	 */
	public static String transBoolean(boolean bl) {
		return bl ? "是" : "否";
	}
	
	/**
	 * 把一个date按照yyyy-MM-dd的格式转化成字符串
	 * 
	 * @param date
	 * @return
	 */
	public static String parseDateToString(Date date){
		return new SimpleDateFormat("yyyy-MM-dd").format(date);
	}
	/**
	 * 将带上comma的一个String转成按comma分割的List<String><br>
	 * 默认的comma是 "," 也可以通过指定comma换成别的符号
	 * 
	 * @param commaString
	 * @param comma
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static List<String> commaStringToStrings(String commaString,String comma) {
		LinkedList<String> strings = new LinkedList<String>();
		for (String string : commaString.split(comma)) {
			strings.add(string);
		}
		return strings;
	}

	/**
	 * 将Collection的每个元素的toString()转成加上comma的一个String<br>
	 * 默认的comma是 "," 也可以通过指定comma换成别的符号
	 * 
	 * @param strings
	 * @param comma
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public static String stringsToCommaString(Collection strings,String comma) {
		String commaString = "";
		if(strings != null)
			for (Object str : strings) 
				commaString += str + comma ;
		return commaString.length() > 0 ? 
				commaString.substring(0,commaString.length() -1) : "　";
	}

	/**
	 * 是否中文
	 */
	public static boolean isCharChinese(char ch) {
		if   (ch >= 0x4e00 && ch <= 0x9fa5)  
			return true;
		return false;
	}

	/**
	 * 是否英文
	 */
	public static boolean isCharEnglish(char ch) {
		if   ((ch >= 0x41 && ch <= 0x5A)
				|| (ch >= 0x61 && ch <= 0x7A))  
			return true;
		return false;
	}

	/**
	 * 是否数字
	 */
	public static boolean isCharNumber(char ch) {
		if   (ch >= 0x30 && ch <= 0x39)  
			return true;
		return false;
	}

	/**
	 * 获得单个汉字的Ascii.
	 * @param cn char 
	 * 汉字字符
	 * @return int
	 * 错误返回 0,否则返回ascii
	 */
	public static int getCnAscii(char cn) {
		byte[] bytes = null;
		try {
			bytes = (String.valueOf(cn)).getBytes("GBK");
		}
		catch (UnsupportedEncodingException ex) {
		}
		if (bytes == null || bytes.length > 2 || bytes.length <= 0) { //错误
			return 0;
		}
		if (bytes.length == 1) { //英文字符
			return bytes[0];
		}
		if (bytes.length == 2) { //中文字符
			int hightByte = 256 + bytes[0];
			int lowByte = 256 + bytes[1];
			int ascii = (256 * hightByte + lowByte) - 256 * 256;
			return ascii;
		}
		return 0; //错误
	}

	/**
	 * 字符串中汉字转为拼音首字母大写
	 * 如果不是汉字则原样输出
	 * @param cn char 
	 * 汉字字符
	 * @return char
	 * 错误返回 '0',否则返回ascii
	 */
	public static String transChineseToFirstPinyin(String str) {
		char[] returnStr = new char[str.length()];
		Map<Character, Integer[]> pinyinMap = getPinyinFirstLetterMap();
		for (int i = 0; i < str.length(); i++) {
			char ch = str.charAt(i);
			//如果该字符不是中文
			if(!isCharChinese(ch)){
				returnStr[i] = ch;
			}else{
				int ascii = getCnAscii(ch);
				//找不到该中文的ascii码
				if(ascii == 0){
					returnStr[i] = ' ';
				}else{
					for (Character key : pinyinMap.keySet()) {
						Integer[] range = pinyinMap.get(key);
						if(ascii >= range[0] && ascii < range[1]){
							returnStr[i] = key.charValue();
							break;
						}
					}
				}
			}
		}
		return String.valueOf(returnStr);
	}

	/**
	 * 查找某个汉字的拼音的首字母的对照表
	 * key是拼音首字母的字符串，value是ascii码范围
	 * 
	 */
	public static Map<Character, Integer[]> getPinyinFirstLetterMap() {
		LinkedHashMap<Character, Integer[]> pinyinMap = new LinkedHashMap<Character, Integer[]>();
		pinyinMap.put('A', new Integer[]{-20319,-20283 });
		pinyinMap.put('B', new Integer[]{-20283,-19775 });
		pinyinMap.put('C', new Integer[]{-19775,-19218 });
		pinyinMap.put('D', new Integer[]{-19218,-18710 });
		pinyinMap.put('E', new Integer[]{-18710, -18526});
		pinyinMap.put('F', new Integer[]{-18526,-18239});
		pinyinMap.put('G', new Integer[]{-18239,-17922 });
		pinyinMap.put('H', new Integer[]{-17922,-17417 });
		pinyinMap.put('J', new Integer[]{-17417,-16474 });
		pinyinMap.put('K', new Integer[]{-16474,-16212 });
		pinyinMap.put('L', new Integer[]{-16212,-15640 });
		pinyinMap.put('M', new Integer[]{-15640,-15165});
		pinyinMap.put('N', new Integer[]{-15165,-14922 });
		pinyinMap.put('O', new Integer[]{-14922, -14914});
		pinyinMap.put('P', new Integer[]{-14914,-14630});
		pinyinMap.put('Q', new Integer[]{-14630,-14149 });
		pinyinMap.put('R', new Integer[]{-14149,-14090 });
		pinyinMap.put('S', new Integer[]{-14090,-13318 });
		pinyinMap.put('T', new Integer[]{-13318,-12838 });
		pinyinMap.put('W', new Integer[]{-12838,-12556});
		pinyinMap.put('X', new Integer[]{-12556,-11847 });
		pinyinMap.put('Y', new Integer[]{-11847,-11055 });
		pinyinMap.put('Z', new Integer[]{-11055,-3246});
		return pinyinMap;
	}
	
}
