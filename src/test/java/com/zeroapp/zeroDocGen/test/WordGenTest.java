package com.zeroapp.zeroDocGen.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import junit.framework.TestCase;

import com.zeroapp.zerodoc.DocMaker;

public class WordGenTest extends TestCase {

	public void testWordGen(){
		try {
			File f = new File("d:\\demo.xml");
			InputStream is = new FileInputStream(f);
			DocMaker docMaker = new DocMaker(is);
			docMaker.saveAllDocToDisk();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
//	public void testConver2PDF(){
//		try{
//			DocMaker.convertDoc2PDF("D:\\temp\\demo.doc");
//		}catch(Exception e){
//			e.printStackTrace();
//		}
//	}

}
