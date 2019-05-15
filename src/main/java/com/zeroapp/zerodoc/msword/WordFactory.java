package com.zeroapp.zerodoc.msword;

import com.jacob2.activeX.ActiveXComponent;
import com.jacob2.com.Dispatch;
import com.jacob2.com.Variant;

/**
 * WordFactory类
 * 
 * @author wkm
 * 
 */
public class WordFactory
{


	/**
	 * word应用程序组件
	 */
	public static ActiveXComponent MsWordApp = null;

	/**
	 * 初始化WordFactory实例
	 * 
	 * @throws Exception
	 */
	public static void openWordService()
	{
		// Open Word if we've not done it already
		if (MsWordApp == null)
		{
			MsWordApp = new ActiveXComponent("Word.Application");
		}
	}

	/**
	 * 打开word应用服务
	 * 
	 * @throws Exception
	 */
	public static void setWordVisible(boolean makeVisible) throws Exception
	{
		// Set the visible property as required.
		Dispatch.put(MsWordApp, "Visible", new Variant(makeVisible));
	}

	/**
	 * 创建新的word实例
	 * 
	 * @throws Exception
	 * 
	 */
	public static WordInstance createNewWordInstance() throws Exception
	{
		// Open Word if we've not done it already
		Dispatch docs = MsWordApp.getProperty("Documents").toDispatch();
		Dispatch doc = Dispatch.call(docs, "Add").toDispatch();

		WordInstance wordInstance = new WordInstance(doc);
		return wordInstance;
	}

	/**
	 * 打开现有的word文档
	 * 
	 * @throws Exception
	 */
	public static WordInstance openWordFile(String word_file) throws Exception
	{
		// Open Word if we've not done it already
		Dispatch docs = MsWordApp.getProperty("Documents").toDispatch();
		Dispatch doc = Dispatch.invoke(docs, "Open", Dispatch.Method,
				new Object[]
				{ word_file }, new int[1]).toDispatch();

		WordInstance wordInstance = new WordInstance(doc);
		return wordInstance;
	}

	/**
	 * 关闭word应用服务
	 */
	public static void closeWordService()
	{
		if (MsWordApp != null)
		{
			Dispatch.call(MsWordApp, "Quit");
			MsWordApp = null;
		}
	}
}
