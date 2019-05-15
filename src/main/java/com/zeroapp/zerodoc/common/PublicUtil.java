package com.zeroapp.zerodoc.common;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.net.URLClassLoader;

import com.zeroapp.zerodoc.exceptions.ZerodocException;

public class PublicUtil
{

	/**
	 * 最大值函数，for long
	 * 
	 * @param a
	 * @param b
	 * @return
	 */
	public static long max(long a, long b)
	{
		return a > b ? a : b;
	}

	/**
	 * 最小值函数，for long
	 * 
	 * @param a
	 * @param b
	 * @return
	 */
	public static long min(long a, long b)
	{
		return a < b ? a : b;
	}

	/**
	 * 最大值函数，for int
	 * 
	 * @param a
	 * @param b
	 * @return
	 */
	public static int max(int a, int b)
	{
		return a > b ? a : b;
	}

	/**
	 * 最小值函数，for int
	 * 
	 * @param a
	 * @param b
	 * @return
	 */
	public static int min(int a, int b)
	{
		return a < b ? a : b;
	}

	/**
	 * 根据文件路径和类名加载对应的类
	 * 
	 * @param file_url
	 * @param class_name
	 * @return
	 * @throws ZerodocException 
	 */
	public static Class loadClass(String file_url, String class_name) throws ZerodocException
			
	{
		Class cls = null;
		URL url = null;

		// 如果未指定路径，则直接在classpath下加载class
		if(file_url == null)
		{
			try{
				cls = Thread.currentThread().getContextClassLoader().loadClass(class_name);
			}catch(Exception e){
				throw new ZerodocException("加载动态类" + class_name + "失败，请确定该类在classpath中。" + e.getMessage());
			}
			
		}
		// 如果指定的是绝对路径，则需要加载文件
		else
		{
			try
			{
				File file = new File(file_url);
				// 如果是文件，获得文件的URL
				if (file.isFile())
				{
					url = file.toURI().toURL();
				}
				// 否则作为普通的URL处理
				else
				{
					url = new URL(file_url);
				}
				
				// 用当前进程所使用的classloader加载类
				URLClassLoader loader = new URLClassLoader(new URL[]
				{ url }, Thread.currentThread().getContextClassLoader());
				
				cls = Class.forName(class_name, true, loader);
			}catch(Exception e)
			{
				throw new ZerodocException("加载动态数据类" + class_name + "失败，请检查设定。file_url=" + file_url + "。" + e.getMessage());
			}
		}

		return cls;
	}

	/**
	 * 根据byte[]生成文件
	 * 
	 * @param b
	 * @param outputFile
	 */
	public static void saveFileFromBytes(byte[] b, String outputFile)
	{
		BufferedOutputStream stream = null;
		try
		{
			FileOutputStream fstream = new FileOutputStream(
					new File(outputFile));
			stream = new BufferedOutputStream(fstream);
			stream.write(b);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		finally
		{
			if (stream != null)
			{
				try
				{
					stream.close();
				}
				catch (IOException e1)
				{
					e1.printStackTrace();
				}
			}
		}
	}

	/**
	 * 将文件读成byte[]
	 * 
	 * @param inputFile
	 * @return
	 */
	public static byte[] getByteFromFile(String inputFile)
	{
		try
		{
			// 开始传输文件
			File file = new File(inputFile);
			byte buffer[] = new byte[(int) file.length()];
			BufferedInputStream input = new BufferedInputStream(
					new FileInputStream(inputFile));
			input.read(buffer, 0, buffer.length);
			input.close();
			return (buffer);
		}
		catch (Exception e)
		{
			System.out.println("FileImpl: " + e.getMessage());
			//e.printStackTrace();
			return (null);
		}

	}

}
