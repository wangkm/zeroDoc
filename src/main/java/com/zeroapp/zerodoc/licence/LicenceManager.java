package com.zeroapp.zerodoc.licence;

import java.io.IOException;

import org.apache.commons.codec.digest.DigestUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.zeroapp.zerodoc.ZeroDocPackage;
import com.zeroapp.zerodoc.util.a;

public class LicenceManager
{

	private static Log logger = LogFactory.getLog(LicenceManager.class);

	public static boolean Validate1(LicenceSetting.LicTypeEnum lic_type,
			String key)
	{

		// 如果是试用版本，则放行
		if (lic_type == LicenceSetting.LicTypeEnum.TRIAL)
		{
			return true;
		}
		else if (key.equals(GenerateKey(lic_type.toString())))
		{
			return true;
		}
		return false;
	}

	public static String GenerateKey(String lic_type)
	{
		String mac_address = "";
		try
		{
			mac_address = a.getMacAddress();
		}
		catch (IOException e)
		{
			logger.fatal("校验序列号出错，请同供应商联系");
			System.exit(-1);
		}
		String md5_1 = DigestUtils.md5Hex(mac_address.toUpperCase()
				+ ZeroDocPackage.get_product_name().toUpperCase());
		String md5_2 = DigestUtils.md5Hex(lic_type.toUpperCase()
				+ ZeroDocPackage.get_version().toUpperCase());
		String md5_3 = DigestUtils.md5Hex(md5_1 + md5_2);

		StringBuffer seed = new StringBuffer("****************");
		for (int i = 0; i < 16; i++)
		{
			seed.setCharAt(i, md5_3.charAt(2 * i + 1));
		}

		char[] aKeyChars =
		{ 105, 88, 64, 90, 86, 50, 74, 78, 88, 37, 72, 51, 79, 38, 71, 53, 67,
				81, 55, 35, 107, 76, 94, 85, 36, 56, 83, 52, 68, 54, 57, 84,
				66, 48, 75, 77 };
		StringBuffer key;
		byte[] keyBytes;
		int patternLength;
		int keyCharsOffset = 0;
		int i = 0;
		int j = 0;

		key = new StringBuffer("#####-#####-#####-#####-#####");
		keyBytes = String.valueOf(seed).getBytes();
		patternLength = key.length();

		while ((i < keyBytes.length) && (j < patternLength))
		{
			keyCharsOffset = keyCharsOffset + Math.abs(keyBytes[i]);
			while (keyCharsOffset >= aKeyChars.length)
			{
				keyCharsOffset = keyCharsOffset - aKeyChars.length;
			}
			while ((key.charAt(j) != 35) && (j < patternLength))
			{
				j++;
			}
			key.setCharAt(j, aKeyChars[keyCharsOffset]);
			if (i == (keyBytes.length - 1))
			{
				i = -1;
			}
			i++;
			j++;
		}
		return key.toString();
	}
}
