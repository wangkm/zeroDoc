package com.zeroapp.zeroDocGen.test;

import java.io.File;
import java.math.BigDecimal;
import java.math.MathContext;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLClassLoader;
import java.text.DecimalFormat;

import com.zeroapp.zerodoc.ZInterface.ZTextProvider;
import com.zeroapp.zerodoc.util.PDFConvertor;

import junit.framework.TestCase;

public class simpleTest extends TestCase{

//	public void testString(){
//		String str1 = null;
//		String str2 = new String();
//		String str3 = new String("");
//		
//		System.out.println(str1.length());
//		System.out.println(str2.length());
//		System.out.println(str3.length());
//	}
//	public void testPrint(){
//		String d = "35465432132165465465";
//		BigDecimal tmp = new BigDecimal(d, new MathContext(10));
//		System.out.println(tmp);
//	}
//	public void testConPdf(){
//		PDFConvertor pdfc = new PDFConvertor();
//		try {
//			pdfc.convert(new String[]{"d:\\三一数字化车间建设方案.doc"});
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//	}
	
//	public void testLoadClass(){
//		File file = new File("D:/Tmp/zeroDE_demo/ext/myProviders.jar");
//		URL url = null;
//		try {
//			url = file.toURI().toURL();
//		} catch (MalformedURLException e) {
//			e.printStackTrace();
//		}
//		URLClassLoader loader = new URLClassLoader(new URL[] { url });
//		
//		Class cls = null;
//		try {
//			cls = Class.forName("demo.myTextProvider", true, loader);
//		} catch (ClassNotFoundException e) {
//			e.printStackTrace();
//		}
//		ZTextProvider provider = null;
//		try {
//			provider = (ZTextProvider)cls.newInstance();
//		} catch (InstantiationException e) {
//			e.printStackTrace();
//		} catch (IllegalAccessException e) {
//			e.printStackTrace();
//		}
//		provider.set_parameters("1, 2, 3, 4, 5");
//		String str = provider.get_value();
//		System.out.println(str);
//	}

}
