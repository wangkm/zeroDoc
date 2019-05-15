package com.zeroapp.zerodoc.util;

import java.net.URL;

/**
 * 将一个可打印的文件转换为PDF格式
 * 
 * @author wkm
 * 
 */
public class PDFConvertor
{
	public void convert(String[] files) throws Exception
	{
		if (files == null)
		{
			System.out.println("请输入待处理的文件");
		}

		URL curr_url = this.getClass().getResource("/Convert2PDF.js");
		String curr_path = curr_url.getFile();
		curr_path = curr_path.substring(1, curr_path.length());
		String cmd = "\"" + curr_path + "\"";

		for (int i = 0; i < files.length; i++)
		{
			cmd = cmd + " \"" + files[i] + "\"";
		}

		cmd = "cmd /c \"" + cmd + "\"";
		Process p = null;
		try
		{
			p = Runtime.getRuntime().exec(cmd);
		}
		catch (Exception e)
		{
			System.out.println("创建进程失败，请在命令行窗口执行以下命令确定路径是否正确：");
			System.out.println(cmd);
		}
		
		int exitValue = -1;
		
		boolean finished = false;
		
		while (!finished)
		{
			try
			{
				exitValue = p.exitValue();
				finished = true;
			}
			catch (Exception e)
			{
				Thread.sleep(100);
			}
		}
		if (exitValue != 0)
		{
			System.out.println("自动调用失败，请在命令行窗口执行以下命令确定路径是否正确：");
			System.out.println(cmd);
		}
		System.out.println("[成功生成pdf文档]");
	}

//	private final int maxTime = 30; // in seconds
//	private final int sleepTime = 250; // in milliseconds
//
//	private int _readyState = 0;
//
//	public void convert2(String[] file_paths) throws InterruptedException
//	{
//		ActiveXComponent fso = null;
//		ActiveXComponent PDFCreator = null;
//
//		fso = new ActiveXComponent("Scripting.FileSystemObject");
//		PDFCreator = new ActiveXComponent("PDFCreator.clsPDFCreator");
//
//		Dispatch.call(PDFCreator, "cStart", "/NoProcessingAtStartup");
//
//		Dispatch cOptions = Dispatch.call(PDFCreator, "cOptions").toDispatch();
//		Dispatch.put(cOptions, "UseAutosave", new Variant(1));
//		Dispatch.put(cOptions, "UseAutosaveDirectory", new Variant(1));
//		Dispatch.put(cOptions, "AutosaveFormat", new Variant(0));
//
//		Variant DefaultPrinter = Dispatch.call(PDFCreator, "cDefaultprinter");
//
//		Dispatch.put(PDFCreator, "cDefaultprinter", new Variant("PDFCreator"));
//
//		Dispatch.call(PDFCreator, "cClearcache");
//
//		for (int i = 0; i < file_paths.length; i++)
//		{
//			// ifname = WScript.arguments.item(i)
//			File file = new File(file_paths[i]);
//
//			if (!file.exists())
//			{
//				System.out.println("Can't find the file: " + file_paths[i]);
//				continue;
//			}
//
//			Boolean isPrintable = Dispatch.call(PDFCreator, "cIsPrintable", file_paths[i]).getBoolean();
//
//			if (!isPrintable)
//			{
//				System.out
//						.println("Converting: "
//								+ file_paths[i]
//								+ "\r\n\r\nAn error is occured: File is not printable!");
//				return;
//			}
//
//			Dispatch.put(cOptions, "AutosaveDirectory", file.getParent());
//			Dispatch.put(cOptions, "AutosaveFilename", new Variant(file_paths[i]));
//			Dispatch.call(PDFCreator, "cPrintfile", new Variant(file_paths[i]));
//			Dispatch.put(PDFCreator, "cPrinterStop", new Variant(false));
//
//			int c = 0;
//			while ((_readyState == 0) && (c < (maxTime * 1000 / sleepTime)))
//			{
//				c = c + 1;
//				Thread.sleep(sleepTime);
//			}
//			if (_readyState == 0)
//			{
//				System.out.println("Converting: " + file_paths[i]
//						+ "\r\n\r\nAn error is occured: Time is up!");
//				break;
//			}
//		}
//
//		Dispatch.put(PDFCreator, "cDefaultprinter", DefaultPrinter);
//		Dispatch.call(PDFCreator, "cClearcache");
//		Thread.sleep(200);
//		Dispatch.call(PDFCreator, "cClose");
//	}

}
