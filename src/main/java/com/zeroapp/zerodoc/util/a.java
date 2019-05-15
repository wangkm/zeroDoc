package com.zeroapp.zerodoc.util;

import java.net.InetAddress;
import java.net.UnknownHostException;
import java.io.InputStream;
import java.io.BufferedInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.StringTokenizer;

public final class a
{

	/**
	 * 获取操作系统类型
	 */
	public final static String getOSType()
	{
		String os_type = System.getProperty("os.name");
		if (os_type.startsWith("Windows"))
		{
			return "Windows";
		}
		else if (os_type.startsWith("Linux"))
		{
			return "Linux";
		}
		else
		{
			return "unknown operating system";
		}
	}

	public final static String getIPAddress() throws UnknownHostException
	{
		return InetAddress.getLocalHost().getHostAddress();
	}

	/**
	 * 获取本机的MAC地址
	 * 
	 * @return
	 * @throws IOException
	 */
	public final static String getMacAddress() throws IOException
	{
		String os = System.getProperty("os.name");

		try
		{
			if (os.startsWith("Windows"))
			{
				return windowsParseMacAddress(windowsRunIpConfigCommand());
			}
			else if (os.startsWith("Linux"))
			{
				return linuxParseMacAddress(linuxRunIfConfigCommand());
			}
			else
			{
				throw new IOException("unknown operating system: " + os);
			}
		}
		catch (ParseException ex)
		{
			ex.printStackTrace();
			throw new IOException(ex.getMessage());
		}
	}

	/*
	 * Linux stuff
	 */
	private final static String linuxParseMacAddress(String ipConfigResponse)
			throws ParseException
	{
		String localHost = null;
		try
		{
			localHost = InetAddress.getLocalHost().getHostAddress();
		}
		catch (java.net.UnknownHostException ex)
		{
			ex.printStackTrace();
			throw new ParseException(ex.getMessage(), 0);
		}

		StringTokenizer tokenizer = new StringTokenizer(ipConfigResponse, "\n");
		String lastMacAddress = null;

		while (tokenizer.hasMoreTokens())
		{
			String line = tokenizer.nextToken().trim();
			boolean containsLocalHost = line.indexOf(localHost) >= 0;

			// see if line contains IP address
			if (containsLocalHost && lastMacAddress != null)
			{
				return lastMacAddress;
			}

			// see if line contains MAC address
			int macAddressPosition = line.indexOf("HWaddr");
			if (macAddressPosition <= 0)
				continue;

			String macAddressCandidate = line.substring(macAddressPosition + 6)
					.trim();
			if (linuxIsMacAddress(macAddressCandidate))
			{
				lastMacAddress = macAddressCandidate;
				continue;
			}
		}

		ParseException ex = new ParseException("cannot read MAC address for "
				+ localHost + " from [" + ipConfigResponse + "]", 0);
		ex.printStackTrace();
		throw ex;
	}

	private final static boolean linuxIsMacAddress(String macAddressCandidate)
	{
		// TODO: use a smart regular expression
		if (macAddressCandidate.length() != 17)
			return false;
		return true;
	}

	private final static String linuxRunIfConfigCommand() throws IOException
	{
		Process p = Runtime.getRuntime().exec("ifconfig");
		InputStream stdoutStream = new BufferedInputStream(p.getInputStream());

		StringBuffer buffer = new StringBuffer();
		for (;;)
		{
			int c = stdoutStream.read();
			if (c == -1)
				break;
			buffer.append((char) c);
		}
		String outputText = buffer.toString();

		stdoutStream.close();

		return outputText;
	}

	/*
	 * Windows stuff
	 */
	private final static String windowsParseMacAddress(String ipConfigResponse)
			throws ParseException
	{
		String localHost = null;
		try
		{
			localHost = InetAddress.getLocalHost().getHostAddress();
		}
		catch (java.net.UnknownHostException ex)
		{
			ex.printStackTrace();
			throw new ParseException(ex.getMessage(), 0);
		}

		StringTokenizer tokenizer = new StringTokenizer(ipConfigResponse, "\n");
		String lastMacAddress = null;

		while (tokenizer.hasMoreTokens())
		{
			String line = tokenizer.nextToken().trim();

			// see if line contains IP address
			if (line.endsWith(localHost) && lastMacAddress != null)
			{
				return lastMacAddress;
			}

			// see if line contains MAC address
			int macAddressPosition = line.indexOf(":");
			if (macAddressPosition <= 0)
				continue;

			String macAddressCandidate = line.substring(macAddressPosition + 1)
					.trim();
			if (windowsIsMacAddress(macAddressCandidate))
			{
				lastMacAddress = macAddressCandidate;
				continue;
			}
		}

		ParseException ex = new ParseException("cannot read MAC address from ["
				+ ipConfigResponse + "]", 0);
		ex.printStackTrace();
		throw ex;
	}

	private final static boolean windowsIsMacAddress(String macAddressCandidate)
	{
		// TODO: use a smart regular expression
		if (macAddressCandidate.length() != 17)
			return false;

		return true;
	}

	private final static String windowsRunIpConfigCommand() throws IOException
	{
		Process p = Runtime.getRuntime().exec("ipconfig /all");
		InputStream stdoutStream = new BufferedInputStream(p.getInputStream());

		StringBuffer buffer = new StringBuffer();
		for (;;)
		{
			int c = stdoutStream.read();
			if (c == -1)
				break;
			buffer.append((char) c);
		}
		String outputText = buffer.toString();

		stdoutStream.close();

		return outputText;
	}
}
