package com.zeroapp.zerodoc;

import java.io.File;
import java.io.InputStream;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import com.zeroapp.zerodoc.msword.WordFactory;
import com.zeroapp.zerodoc.util.DocTypeEnum;
import com.zeroapp.zerodoc.util.PDFConvertor;

/**
 * 根据XML定义来生成标准化文档的类 可以一次生成多个文档。
 * 
 * @author wkm
 * 
 */
public class DocMaker
{

	// 指定待生成的文档类型，如word、pdf等等
	private DocTypeEnum _doc_type = null;

	// 文档信息的xml对象
	private Document _doc_info = null;
	
	// 文档对象
	private ZeroDocument _zerodoc = null;

	// 缺省的文档保存路径
	private String _default_save_path = null;


	
	/**
	 * 构造函数，直接接收XML Document
	 * @param xmlDoc
	 * @throws Exception 
	 */
	public DocMaker(Document xmlDoc) throws Exception{
		set_doc_info(xmlDoc);
		set_doc_type(DocTypeEnum.MSWord);

		try{
			init();
		}catch(Exception e){
			System.out.println(e.getMessage());
			throw e;
		}		
	}


	/**
	 * 构造函数，指定文件路径
	 * 
	 * @param doc_definition
	 * @throws Exception
	 */
	public DocMaker(String doc_path) throws Exception
	{
		File file = new File(doc_path);
		if(!file.exists() || !file.canRead()){
			throw new Exception("指定的文件：" + doc_path + "不存在或不可访问");
		}
		SAXReader reader = new SAXReader();
		set_doc_info(reader.read(file));
		set_doc_type(DocTypeEnum.MSWord);

		try{
			init();
		}catch(Exception e){
			System.out.println(e.getMessage());
			throw e;
		}		
	}


	/**
	 * 构造函数，接收inputStream流
	 * 
	 * @param xml_ins,
	 *            xml流
	 * @throws Exception
	 */
	public DocMaker(InputStream xml_ins) throws Exception
	{
		SAXReader reader = new SAXReader();
		set_doc_info(reader.read(xml_ins));
		set_doc_type(DocTypeEnum.MSWord);

		try{
			init();
		}catch(Exception e){
			System.out.println(e.getMessage());
			throw e;
		}
	}

	/**
	 * 初始化函数
	 * 
	 * @param doc_definition
	 * @param word
	 * @throws Exception
	 */
	private void init() throws Exception
	{
		// ---------- start of 根节点属性 ---------
		Element root = _doc_info.getRootElement();
		// ========== end of 根节点属性 ===========
		
		_zerodoc = new ZeroDocument(root);

	}

	/**
	 * create and save all of the documents to the disk
	 * 
	 * @throws Exception
	 */
	public void saveAllDocToDisk() throws Exception
	{
		if (_doc_type != DocTypeEnum.MSWord && _doc_type != DocTypeEnum.PDF)
		{
			System.out.println("非法的文档类型。");
			return;
		}
		try{
			// 启动word服务
			WordFactory.openWordService();
			
			// 生成word文档
			_zerodoc.SaveDocToDisk();
			
			// 关闭word服务
			WordFactory.closeWordService();
		}catch(Exception e){
			System.out.println(e.getMessage());
			// 关闭word服务
			WordFactory.closeWordService();
			// 抛出异常
			throw e;
		}
	}
	
	/**
	 *  将指定的word文档转换为pdf文档
	 *  生成的pdf文档与word文档同名，且在同一个目录下
	 * @param docFile 待处理的word文档全路径名，如 "D:\\tmp\\demo.doc"
	 * @throws Exception
	 */
	public static void convertDoc2PDF(String docFile) throws Exception{
		PDFConvertor convertor = new PDFConvertor();
		convertor.convert(new String[]{docFile});
	}

	/**
	 * get documents' type
	 * 
	 * @return
	 */
	public DocTypeEnum get_doc_type()
	{
		return _doc_type;
	}

	/**
	 * set documents' type
	 * 
	 * @param _doc_type
	 */
	private void set_doc_type(DocTypeEnum _doc_type)
	{
		this._doc_type = _doc_type;
	}

	/**
	 * get documents list
	 * 
	 * @return
	 */

	/**
	 * get the xml information of document
	 * 
	 * @return
	 */
	public Document get_doc_info()
	{
		return _doc_info;
	}

	/**
	 * set the xml information of document
	 * 
	 * @param _doc_info
	 */
	private void set_doc_info(Document _doc_info)
	{
		this._doc_info = _doc_info;
	}

	/**
	 * get the default save path
	 * 
	 * @return
	 */
	public String get_default_save_path()
	{
		return _default_save_path;
	}

	/**
	 * set the default save path
	 * 
	 * @param _default_sava_path
	 */
	public void set_default_save_path(String _default_save_path)
	{
		this._default_save_path = _default_save_path;
	}

}
