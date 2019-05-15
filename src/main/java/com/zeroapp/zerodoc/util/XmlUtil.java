package com.zeroapp.zerodoc.util;


import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringWriter;

import org.dom4j.Document;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

public class XmlUtil {
	public static final String GB2312 = "GB2312";
	public static final String UTF8 = "UTF-8";

	/**
	 * 产生dom树
	 */
	public static Document generateDocFromInputStream(InputStream ins) {
		Document document = null;
		SAXReader saxReader = new SAXReader();
		byte[] bytes;
		try {
			bytes = IoUtil.readBytes(ins);
		} catch (IOException e) {
			throw new RuntimeException("io 异常.");
		}
		try {
			String strTmp = new String(bytes, "UTF-8").trim();
			document = saxReader.read(new ByteArrayInputStream(strTmp
					.getBytes("UTF-8")));
			document.setXMLEncoding("UTF-8");
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
		return document;
	}

	public static Document generateDocFromInputStream(InputStream ins,
			String encode) {
		Document document = null;
		SAXReader saxReader = new SAXReader();
		byte[] bytes;
		try {
			bytes = IoUtil.readBytes(ins);
		} catch (IOException e) {
			throw new RuntimeException("io 异常.");
		}
		try {
			String strTmp = new String(bytes, encode).trim();
			document = saxReader.read(new ByteArrayInputStream(strTmp
					.getBytes(encode)));
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
		return document;
	}

	public static Document generateDocFromInputStream(InputStream ins,
			String fromCode, String toCode) {
		Document document = null;
		SAXReader saxReader = new SAXReader();
		byte[] bytes;
		try {
			bytes = IoUtil.readBytes(ins);
		} catch (IOException e) {
			throw new RuntimeException("io 异常.");
		}
		try {
			String strTmp = new String(bytes, fromCode).trim();
			document = saxReader.read(new ByteArrayInputStream(strTmp
					.getBytes(toCode)));
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage());
		}
		return document;
	}
	
	/**
	 * 将document保存为文件
	 * @param document
	 * @param file
	 */
	public static void saveXmlToFile(Document document, File file) {
		try {			
			XMLWriter xmlWriter = new XMLWriter(new FileOutputStream(file));
			xmlWriter.write(document);
			xmlWriter.close();
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage());
		} finally {
			// do nothing
		}
	}
	public static String saveXmlToString(Document document) {
		try {
			StringWriter stringWriter = new StringWriter();
			OutputFormat format = OutputFormat.createCompactFormat();
            format.setEncoding("utf-8"); 

			XMLWriter xmlWriter = new XMLWriter(stringWriter, format);
			xmlWriter.write(document);
			xmlWriter.close();
			return stringWriter.toString();
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage());
		}
	}

}
