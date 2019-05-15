package com.zeroapp.zerodoc;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.dom4j.Element;

import com.zeroapp.zerodoc.ZInterface.ZImageProvider;
import com.zeroapp.zerodoc.ZInterface.ZTableRangeProvider;
import com.zeroapp.zerodoc.ZInterface.ZTextProvider;
import com.zeroapp.zerodoc.common.Cell_definition;
import com.zeroapp.zerodoc.common.Data_unit;
import com.zeroapp.zerodoc.common.Merge_definition;
import com.zeroapp.zerodoc.common.Parameter;
import com.zeroapp.zerodoc.common.PublicUtil;
import com.zeroapp.zerodoc.common.TableRange_data;
import com.zeroapp.zerodoc.exceptions.ZerodocException;
import com.zeroapp.zerodoc.msword.Cell;
import com.zeroapp.zerodoc.msword.Image;
import com.zeroapp.zerodoc.msword.Range;
import com.zeroapp.zerodoc.msword.Section;
import com.zeroapp.zerodoc.msword.Table;
import com.zeroapp.zerodoc.msword.WordFactory;
import com.zeroapp.zerodoc.msword.WordInstance;
import com.zeroapp.zerodoc.msword.util.AlignmentType;
import com.zeroapp.zerodoc.msword.util.AutoFitBehavior;
import com.zeroapp.zerodoc.msword.util.BorderStyle;
import com.zeroapp.zerodoc.msword.util.ConstantUtil;
import com.zeroapp.zerodoc.msword.util.FontStyle;
import com.zeroapp.zerodoc.msword.util.ParaStyle;

/**
 * zeroDoc的基本类 每个实例代表一份文档 可以利用SaveDocToDisk函数将文档保存为本地文件
 * 
 * @author wkm
 * 
 */
public class ZeroDocument {
	// 文档实例
	WordInstance _word_instance = null;

	// 文档的保存目录
	private String _save_path = null;

	// 文档名称
	private String _name = null;

	// 文档的全路径
	String _file_path_name = null;

	// 如果存在同名文件，是否覆盖
	private boolean _overwrite = true;

	// 文档title
	private String _title = null;

	// 文档创建者
	private String _creator = null;

	// 文档创建时间
	private String _create_time = null;

	// 文档最后修改者
	private String _last_modifier = null;

	// 文档最后修改时间
	private String _last_modtime = null;

	// 文档备注
	private String _remark = null;

	// 文档的描述信息
	private Element _xml_info = null;

	// 记录当前已经处理的节点数目
	private int _dealed_element = 0;

	// 记录当前许可证可以支持的最大节点数目
	private int _max_element = 10000;

	/**
	 * 构造函数
	 * 
	 * @param xml_info
	 * @throws Exception
	 * @throws Exception
	 */
	public ZeroDocument(Element xml_info) throws Exception {

		// 初始化已经处理的节点数目
		_dealed_element = 0;

		// 符合许可证要求，为_xml_info和_save_path赋值
		_xml_info = xml_info;

		// a template element
		Element element_tmp = null;

		Element meta_info = _xml_info.element("metadata");
		if (meta_info == null) {
			throw new Exception("错误：未指定文档的metadata");
		}

		_title = meta_info.attributeValue("title");
		_creator = meta_info.attributeValue("creator");
		_create_time = meta_info.attributeValue("create_time");
		_last_modifier = meta_info.attributeValue("last_modifier");
		_last_modtime = meta_info.attributeValue("last_modtime");
		_remark = meta_info.attributeValue("remark");

		String save_path = meta_info.attributeValue("save_path");

		// 如果指定了本文档的保存路径，则覆盖缺省路径
		if (save_path == null) {
			throw new Exception("请指定文档的保存目录");
		} else {
			File testSavePath = new File(save_path);
			if (testSavePath.isDirectory()) {
				_save_path = testSavePath.getPath();
			} else {
				throw new Exception("当前文档的保存目录'" + save_path + "'有误");
			}
		}

		// set file name
		_name = meta_info.attributeValue("file_name");

		// get the real path name for document to save
		_file_path_name = _save_path + "\\" + _name;

		// should overwrite or not
		_overwrite = meta_info.attributeValue("overwrite").equalsIgnoreCase(
				"true") ? true : false;

	}

	/**
	 * 将文档保存到本地 目前支持word类型的文档
	 * 
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
	public void SaveDocToDisk() throws Exception {
		System.out
				.println("*************** ZeroDoc Generator *********************");
		System.out.println("物理文件路径：\t" + this.get_file_path_name());
		System.out.println("覆盖已有文件：\t" + this.is_overwrite());

		if (this.get_title() != null)
			System.out.println("文档标题：\t" + this.get_title());

		if (this.get_creator() != null)
			System.out.println("文档创建者：\t" + this.get_creator());

		if (this.get_create_time() != null)
			System.out.println("文档创建时间：\t" + this.get_create_time());

		if (this.get_last_modifier() != null)
			System.out.println("最后修改者：\t" + this.get_last_modifier());

		if (this.get_last_modtime() != null)
			System.out.println("最后修改时间：\t" + this.get_last_modtime());

		if (this.get_remark() != null)
			System.out.println("文档备注信息：\t" + this.get_remark());
		System.out
				.println("-------------------------------------------------------");

		// 如果选择不覆盖，但目标文件已经存在，则提示
		if (!_overwrite) {
			File f = new File(_file_path_name);
			if (f.exists()) {
				throw new Exception("目标文件'" + _file_path_name
						+ "'已经存在。请选择别的文件名，或修改overwrite为true。");
			}
		}
		try {
			// 创建新的word实例
			_word_instance = WordFactory.createNewWordInstance();

			// 开始处理文档中的各个数据节
			List<Element> sections = _xml_info.elements("section");

			// deal with all of current sections
			for (int i = 0; i < sections.size(); i++) {

				Element s = sections.get(i);

				// 新建一个section（如果是第一个，则不用新建）
				Section section = null;
				if (i == 0) {
					section = _word_instance.getLastSection();
				} else {
					String section_type = s.attributeValue("section_type");
					section = _word_instance.addSection(section_type);
				}

				// 本节的页面设置
				section.setPageProperties(s);

				// get the elements of current section
				List<Element> elements = s.elements();

				// deal with the elements
				for (Element element : elements) {
					try {
						// 将光标移动到文档最后
						_word_instance.moveCursorToEnd();
						// 添加单个节点
						createElement(element);
					} catch (ZerodocException e) {
						// 如果是Zerodoc自己的异常，不影响文档生成。
						System.out.println("出现异常：" + e.getMessage());
					} catch (Exception e) {
						throw e;
					}
				}
			}

			// ----------------------------------------------
			// 更新目录
			// ----------------------------------------------
			_word_instance.updateContent();

			// ----------------------------------------------
			// 保存文档
			// ----------------------------------------------
			try {
				_word_instance.saveas(_file_path_name);
			} catch (Exception e) {
				throw new ZerodocException("保存文件失败，请确定文件的保存路径'"
						+ _file_path_name + "'是合法的文件路径名，且同名文件没有处于只读状态。");
			}

			// ----------------------------------------------
			// 关闭word文档
			// ----------------------------------------------
			_word_instance.close(ConstantUtil.wdDoNotSaveChanges);

			// ----------------------------------------------
			// print the result
			// ----------------------------------------------
			System.out.println("[创建完毕]");
		} catch (Exception e) {
			// ----------------------------------------------
			// 关闭word文档
			// ----------------------------------------------
			_word_instance.close(ConstantUtil.wdDoNotSaveChanges);

			// print the result
			System.out.println("[创建失败]");
			e.printStackTrace();
			throw new Exception(e.getMessage());
		}
	}

	/**
	 * 创建页面元素的主函数 根据元素类型调用不同的创建函数
	 * 
	 * @throws Exception
	 */
	private void createElement(Element element) throws Exception {
		if (element == null)
			return;

		// 验证待处理文档的element数目是否超出许可证允许的最大数目
		if (++_dealed_element > _max_element) {
			throw new Exception("当前文档的element节点数（" + _dealed_element
					+ "）超出许可证允许的最大节点数（" + _max_element + "），请同供应商联系。");
		}

		String element_type = element.getName();

		System.out.println("[" + _dealed_element + "]正在处理" + element_type
				+ "...");

		try{
		// 页眉
		if (element_type.equalsIgnoreCase("page_header")) {
			createPageHeader(element);
		}

		// 页眉
		else if (element_type.equalsIgnoreCase("page_footer")) {
			createPageFooter(element);
		}
		// 硬回车
		else if (element_type.equalsIgnoreCase("hard_return")) {
			createHardReturn(element);
		}

		// 软回车
		else if (element_type.equalsIgnoreCase("soft_return")) {
			createSoftReturn(element);
		}

		// 分页符
		else if (element_type.equalsIgnoreCase("new_page")) {
			createNewPage(element);
		}

		// 静态文本
		else if (element_type.equalsIgnoreCase("static_text")) {
			createStaticText(element);
		}

		// 动态文本
		else if (element_type.equalsIgnoreCase("dynamic_text")) {
			createDynamicText(element);
		}

		// 静态图片
		else if (element_type.equalsIgnoreCase("static_image")) {
			createStaticImage(element);
		}

		// 动态图片
		else if (element_type.equalsIgnoreCase("dynamic_image")) {
			createDynamicImage(element);
		}

		// 目录
		else if (element_type.equalsIgnoreCase("content")) {
			createContent(element);
		}

		// 表格
		else if (element_type.equalsIgnoreCase("table")) {
			createTable(element);
		}

		// 制表符
		else if (element_type.equalsIgnoreCase("tab")) {
			createTab(element);
		}

		// 当前页码
		else if (element_type.equalsIgnoreCase("page_number")) {
			createPageNumber(element);
		}

		// 当前页数
		else if (element_type.equalsIgnoreCase("current_page_count")) {
			createCurrentPageCount(element);
		}

		// 总页数
		else if (element_type.equalsIgnoreCase("page_count")) {
			createPageCount(element);
		}

		// 日期时间
		// -----------------------------------------------
		else if (element_type.equalsIgnoreCase("dateTime_stamp")) {
			createDateTimeStamp(element);
		}
		// 样式，忽略
		else if (element_type.equalsIgnoreCase("style")) {
			// do nothing
		}
		// unsupported type
		else {
			// throw an exception to the caller
			throw new ZerodocException("不支持的类型: " + element_type);
		}
		}catch(Exception e){
			throw new Exception("未能成功创建" + element_type + "：" + e.getMessage());
		}
	}

	/**
	 * get the save path of the document
	 * 
	 * @return
	 */
	public String get_save_path() {
		return _save_path;
	}

	/**
	 * set the save path of the document
	 * 
	 * @return
	 */
	public void set_save_path(String _save_path) {
		this._save_path = _save_path;
	}

	/**
	 * set the document's name
	 * 
	 * @return
	 */
	public String get_name() {
		return _name;
	}

	/**
	 * get the document's name
	 * 
	 * @param _name
	 */
	public void set_name(String _name) {
		this._name = _name;
	}

	/**
	 * set if overwrite existed file
	 * 
	 * @param _overwrite
	 */
	public void set_overwrite(boolean _overwrite) {
		this._overwrite = _overwrite;
	}

	/**
	 * get if overwrite existed file
	 * 
	 * @return
	 */
	public boolean is_overwrite() {
		return _overwrite;
	}

	/**
	 * get the xml information
	 * 
	 * @return
	 */
	public Element get_xml_info() {
		return _xml_info;
	}

	/**
	 * set the xml information
	 * 
	 * @param _xml_info
	 */
	public void set_xml_info(Element _xml_info) {
		this._xml_info = _xml_info;
	}

	/**
	 * 生成页眉
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createPageHeader(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("page_header"));

		Section section = _word_instance.getLastSection();

		// 设定最后一节的页眉不链接到前一节
		section.setHeaderLinkToPrevious(false);
		section.clearHeader();
		// 根据设定去掉页眉的横线
		String with_line = element.attributeValue("with_line");
		if ("false".equalsIgnoreCase(with_line)) {
			section.removeHeaderLine();
		}

		_word_instance.seekToPageHeader();

		// get the elements of current section
		List<Element> elements = element.elements();

		// deal with the elements
		for (Element e : elements) {
			try {
				// 将光标移动到当前选中区域的末尾
				_word_instance.moveCursorToEnd();
				// 处理单个节点
				createElement(e);
			} catch (ZerodocException ex) {
				// 如果是Zerodoc自己的异常，不影响文档生成。
				System.out.println("出现异常：" + ex.getMessage());
			} catch (Exception ex) {
				throw ex;
			}
		}

		// 返回到文档正文
		_word_instance.seekToMainDocument();
	}

	/**
	 * 生成页脚
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createPageFooter(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("page_footer"));

		// 设定最后一节的页眉不链接到前一节
		Section section = _word_instance.getLastSection();
		section.setFootLinkToPrevious(false);
		section.clearFooter();
		_word_instance.seekToPageFooter();

		// get the elements of current section
		List<Element> elements = element.elements();

		// deal with the elements
		for (Element e : elements) {
			try {
				// 将光标移动到当前选中区域的末尾
				_word_instance.moveCursorToEnd();
				// 处理单个节点
				createElement(e);
			} catch (ZerodocException ex) {
				// 如果是Zerodoc自己的异常，不影响文档生成。
				System.out.println("出现异常：" + ex.getMessage());
			} catch (Exception ex) {
				throw ex;
			}
		}

		// 返回到文档正文
		_word_instance.seekToMainDocument();
	}

	/**
	 * 生成硬回车
	 */
	private void createHardReturn(Element element) {
		assert (element.getName().equalsIgnoreCase("hard_return"));

		_word_instance.newLine();
	}

	/**
	 * 生成软回车
	 */
	private void createSoftReturn(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("soft_return"));

		char[] a = { 11 };
		String strTmp = new String(a);
		_word_instance.appendText(strTmp, null);

	}

	/**
	 * 生成制表符
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createTab(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("tab"));

		// use \t to implement the tab
		_word_instance.appendText("\t", null);

	}

	/**
	 * 生成目录
	 * 
	 * @param element
	 */
	private void createContent(Element element) {
		assert (element.getName().equalsIgnoreCase("content"));
		_word_instance.addContent();
	}

	/**
	 * 生成页码
	 * 
	 * @param element
	 */
	private void createPageNumber(Element element) {
		assert (element.getName().equalsIgnoreCase("page_number"));
		_word_instance.addPageNumber();

	}

	/**
	 * 生成文档总页数
	 * 
	 * @param element
	 */
	private void createPageCount(Element element) {
		assert (element.getName().equalsIgnoreCase("page_count"));
		_word_instance.addPageCount();
	}

	/**
	 * 生成当前的总页数
	 * 
	 * @param element
	 */
	private void createCurrentPageCount(Element element) {
		assert (element.getName().equalsIgnoreCase("current_page_count"));

		String offsetStr = element.attributeValue("offset");
		int offset;
		try {
			offset = Integer.parseInt(offsetStr);
		} catch (Exception e) {
			offset = 0;
		}
		try {
			_word_instance.addCurrentPageCount(offset);
		} catch (ZerodocException e) {
			System.out.println("出现异常：" + e.getMessage());
		}
	}

	/**
	 * 生成日期时间标签
	 * 
	 * @param element
	 */
	private void createDateTimeStamp(Element element) {
		assert (element.getName().equalsIgnoreCase("date_stamp"));

		String format = element.attributeValue("dateTimeFormat");
		try {
			boolean InsertAsField = Boolean.parseBoolean(element
					.attributeValue("insertAsField"));
			_word_instance.addDataTime_stamp(format, InsertAsField);
		} catch (Exception e) {
			System.out.println("创建日期时间失败：" + e.getMessage());
			return;
		}

	}

	/**
	 * 生成分页符
	 * 
	 * @param element
	 */
	private void createNewPage(Element element) {
		assert (element.getName().equalsIgnoreCase("new_page"));
		_word_instance.newPage();
	}

	/**
	 * 生成静态文本
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createStaticText(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("text"));

		// _word_instance.moveToEnd();
		Element text = element;

		// 当前文本的样式
		Element style_info = text.element("style");

		String text_detail = text.elementText("static_data");
		_word_instance.appendText(text_detail, style_info);
	}

	/**
	 * 生成动态文本
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createDynamicText(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("text"));

		Element text = element;

		// 当前文本的样式
		Element style_info = text.element("style");

		Element provider = text.element("data_source").element("provider");
		String file_url = provider.attributeValue("file_url");
		String class_name = provider.attributeValue("class_name");
		String options = provider.attributeValue("options");
		Element parameters = text.element("data_source").element("parameters");

		ZTextProvider textProvider = null;

		// try to load the user defined class
		try {
			Class cls = PublicUtil.loadClass(file_url, class_name);
			textProvider = (ZTextProvider) cls.newInstance();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return;
		}

		// 传递特殊设定
		textProvider.set_options(options);

		// 获取并传递参数列表
		if (parameters != null) {
			List<Parameter> para_list = new ArrayList();
			for (Element p : (List<Element>) parameters.elements("paramater")) {
				String name = p.attributeValue("name");
				String value = p.attributeValue("value");
				String remark = p.attributeValue("remark");
				Parameter para = new Parameter(name, value, remark);
				para_list.add(para);
			}

			// set the parameter of the user defined
			// class
			textProvider.set_parameters(para_list);
		}

		// 获取结果
		List<Data_unit> data_list = null;
		try {
			data_list = textProvider.get_text();
		} catch (Exception e) {
			System.out.println("获取动态文本数据失败，可能是接口不匹配：" + e.getMessage());
		}

		if (data_list == null) {
			return;
		}
		for (Data_unit data_unit : data_list) {
			// 对文本插件，只支持文本类型，且只设定文字样式，不设定段落样式
			if (data_unit.get_type().equalsIgnoreCase("text")
					&& data_unit.get_data() instanceof String) {

				// add the text with certain
				// paragraph and font style
				_word_instance.appendText((String) data_unit.get_data(),
						style_info);
			} else {
				throw new Exception("非法的类型:" + data_unit.get_type());
			}
		}

	}

	/**
	 * 生成静态图片
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createStaticImage(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("image"));

		// move the cursor the the end to avoid overwrite

		Element image = element;

		Image image_obj = null;

		Range range = _word_instance.getCurrentRange();
		range.setDefaultStyle();
		String file_url = image.attributeValue("file_url");

		try {
			image_obj = _word_instance.addImage(file_url);
		} catch (Exception e) {
			// 创建失败，则打印消息并返回
			System.out.println("未能创建静态图片：" + e.getMessage());
			return;
		}

		// 设定当前图片的样式
		Element style_info = image.element("style");
		image_obj.setStyle(style_info);
	}

	/**
	 * 生成动态图片
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createDynamicImage(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("image"));

		// move the cursor the the end to avoid overwrite

		Element image = element;

		Image image_obj = null;

		Element provider = image.element("data_source").element("provider");

		String file_url = provider.attributeValue("file_url");

		String class_name = provider.attributeValue("class_name");

		String options = provider.attributeValue("options");

		Element parameters = image.element("data_source").element("parameters");

		ZImageProvider imageProvider = null;

		try {
			Class cls = PublicUtil.loadClass(file_url, class_name);
			imageProvider = (ZImageProvider) cls.newInstance();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return;
		}

		// 传递特殊设定
		imageProvider.set_options(options);

		// 获取并传递参数列表
		if (parameters != null) {
			List<Parameter> para_list = new ArrayList();
			for (Element p : (List<Element>) parameters.elements("paramater")) {
				String name = p.attributeValue("name");
				String value = p.attributeValue("value");
				String remark = p.attributeValue("remark");
				Parameter para = new Parameter(name, value, remark);
				para_list.add(para);
			}
			imageProvider.set_parameters(para_list);
		}

		// 获取结果
		Object image_data = null;
		try {
			image_data = imageProvider.get_image();
		} catch (Exception e) {
			System.out.println("获取动态图片数据失败，可能是接口不匹配：" + e.getMessage());
		}

		if (image_data == null) {
			return;
		}

		// 如果是文件路径
		if (image_data instanceof String) {
			image_obj = _word_instance.addImage((String) image_data);
		}
		// 如果是文件数据
		else if (image_data instanceof byte[]) {
			File file = File.createTempFile("~zerodoc", ".tmp");
			file.deleteOnExit();
			PublicUtil.saveFileFromBytes((byte[]) image_data, file.getPath());
			image_obj = _word_instance.addImage(file.getPath());
		}
		// 其它类型，抛出异常
		else {
			throw new ZerodocException("未知或非法的动态图片返回数据：" + image_data);
		}
		imageProvider = null;
		if (image_obj == null) {
			throw new ZerodocException("生成动态图片失败。");
		}

		// 设定当前图片的样式
		Element style_info = image.element("style");
		image_obj.setStyle(style_info);
	}

	/**
	 * 生成表格
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createTable(Element element) throws Exception {
		assert (element.getName().equalsIgnoreCase("table"));

		Element table = element;

		// 获取表格初始的行数和列数
		String table_rows_s = table.attributeValue("rows");
		String table_columns_s = table.attributeValue("columns");

		int table_rows;
		int table_columns;
		try {
			table_rows = Integer.parseInt(table_rows_s);
			table_columns = Integer.parseInt(table_columns_s);
		} catch (Exception e) {
			System.out.println("解析表格的行列出错，请确认。row=" + table_rows_s
					+ ", column=" + table_columns_s + "." + e.getMessage());
			return;
		}
		// 创建表格
		Table table_obj = _word_instance.addTable(table_rows, table_columns);

		// 清除表格样式
		table_obj.clearFormatting();

		// --- 设定整个表格的样式 --------------------------------------------------
		Element table_style_info = table.element("style");

		if (table_style_info != null) {
			// 对齐方式
			String table_alignment = table_style_info
					.attributeValue("alignment");

			if (table_alignment != null) {
				table_obj.setAlignment(AlignmentType
						.getAlignmentID(table_alignment));
			}

			// 自动调整机制
			String auto_fit = table_style_info.attributeValue("autofit");

			if (auto_fit != null) {
				table_obj.setAutoFitBehavior(AutoFitBehavior
						.getAutoFitID(auto_fit));
			}

			// 设定表格边框
			BorderStyle tableBorderStyle = new BorderStyle(table_style_info);
			table_obj.setBorderStyle(tableBorderStyle);

		}
		// 设定表格中每个单元格的样式.
		// 由于对每个单元格而言，是继承父对象的样式，所以边框属性无效
		for (int i = 1; i <= table_rows; i++) {
			for (int j = 1; j <= table_columns; j++) {
				Cell cell_obj = table_obj.getCell(i, j);
				cell_obj.setStyle(table_style_info, true);
			}
		}
		// ======================================================================

		// --- 处理各个Range ------------------------------------------------------
		Element ranges = table.element("ranges");
		if (ranges != null) {
			List<Element> range_list = ranges.elements();

			for (Element range : range_list) {
				String range_type = range.getName();
				if (range_type.equalsIgnoreCase("static_range")) {
					createStaticRange(table_obj, range);
				} else if (range_type.equalsIgnoreCase("dynamic_range")) {
					createDynamicRange(table_obj, range);
				} else {
					throw new Exception("非法的Range类型：" + range_type);
				}
			}
		}
		// ======================================================================
	}

	/**
	 * 创建表格中的静态区域
	 * 
	 * @param element
	 * @throws Exception
	 */
	private void createStaticRange(Table table_obj, Element element)
			throws Exception {
		if (table_obj == null || element == null) {
			return;
		}

		assert (element.getName().equalsIgnoreCase("static_range"));

		// 区域位置
		String start_row_s = element.attributeValue("start_row");
		String start_column_s = element.attributeValue("start_column");
		String end_row_s = element.attributeValue("end_row");
		String end_column_s = element.attributeValue("end_column");

		int start_row;
		int start_column;
		int end_row;
		int end_column;
		try {
			start_row = PublicUtil.min(Integer.parseInt(start_row_s), Integer
					.parseInt(end_row_s));
			start_column = PublicUtil.min(Integer.parseInt(start_column_s),
					Integer.parseInt(end_column_s));
			end_row = PublicUtil.max(Integer.parseInt(start_row_s), Integer
					.parseInt(end_row_s));
			end_column = PublicUtil.max(Integer.parseInt(start_column_s),
					Integer.parseInt(end_column_s));
		} catch (Exception e) {
			System.out.println("解析动态区域的行列设定出错，请确认。" + e.getMessage());
			return;
		}

		// 区域样式，目前只支持边框设定
		Element range_style_info = element.element("style");
		BorderStyle range_borderStyle = new BorderStyle(range_style_info);

		table_obj.SetRangeBorderStyle(range_borderStyle, start_row,
				start_column, end_row, end_column);

		// 设定当前区域中的所有单元格的属性
		// 由于对每个单元格而言，是继承父对象的样式，所以边框属性无效
		for (int i = start_row; i <= end_row; i++) {
			for (int j = start_column; j <= end_column; j++) {
				Cell cell_obj = table_obj.getCell(i, j);
				cell_obj.setStyle(range_style_info, true);
			}
		}

		// 处理静态区域中的cell
		Element cells = element.element("cells");
		if (cells != null) {
			List<Element> cell_list = cells.elements("cell");

			for (Element cell : cell_list) {
				// 单元格位置，Range中的单元格均为相对位置，需要转换为绝对位置
				String row = cell.attributeValue("row");
				String column = cell.attributeValue("column");
				try {
					Integer.parseInt(row);
					Integer.parseInt(column);
				} catch (Exception e) {
					System.out.println("解析单元格的位置出错，请确认。row=" + row
							+ ", column=" + column + "." + e.getMessage());
				}

				// 获取当前cell在表格中的绝对坐标
				int absolut_row = start_row + Integer.parseInt(row) - 1;
				int absolut_column = start_column + Integer.parseInt(column)
						- 1;

				// 目标row比当前表格已有的row大，则新建row
				if (absolut_row > table_obj.getRowCount()) {
					table_obj.addRow(absolut_row - table_obj.getRowCount());
				}

				// 目标row比当前表格已有的row大，则新建row
				if (absolut_column > table_obj.getColumnCount()) {
					table_obj.addColumn(absolut_column
							- table_obj.getColumnCount());
				}

				Cell cell_obj = table_obj.getCell(absolut_row, absolut_column);

				// 生成单元格内部的内容
				cell_obj.select();
				createStaticCell(cell_obj, cell);
			}
		}

		// 处理Range中需要合并的表格
		Element merge_infos = element.element("merge_infos");
		table_obj.batchMerge(merge_infos);
	}

	/**
	 * 可以在单元格中插入任意类型的数据
	 * 
	 * @param cell_obj
	 * @param element
	 * @throws Exception
	 */
	private void createStaticCell(Cell cell_obj, Element element)
			throws Exception {
		// get the elements of current section
		List<Element> elements = element.elements();

		// 设定单元格的样式
		Element cell_style_info = element.element("style");
		cell_obj.setStyle(cell_style_info, false);

		// deal with the elements
		for (Element e : elements) {
			try {
				// 将光标移动到cell的末尾
				cell_obj.moveCursorToEnd();
				// 处理单个节点
				createElement(e);
			} catch (ZerodocException ex) {
				// 如果是Zerodoc自己的异常，不影响文档生成。
				System.out.println("出现异常：" + ex.getMessage());
			} catch (Exception ex) {
				throw ex;
			}
		}

	}

	/*
	 * 创建表格中的动态区域
	 */
	private void createDynamicRange(Table table_obj, Element element)
			throws Exception {
		if (table_obj == null || element == null) {
			return;
		}

		assert (element.getName().equalsIgnoreCase("dynamic_range"));

		// 单元格位置
		String start_row_s = element.attributeValue("start_row");
		String start_column_s = element.attributeValue("start_column");
		String end_row_s = element.attributeValue("end_row");
		String end_column_s = element.attributeValue("end_column");

		int start_row;
		int start_column;
		int end_row;
		int end_column;
		try {
			start_row = PublicUtil.min(Integer.parseInt(start_row_s), Integer
					.parseInt(end_row_s));
			start_column = PublicUtil.min(Integer.parseInt(start_column_s),
					Integer.parseInt(end_column_s));
			end_row = PublicUtil.max(Integer.parseInt(start_row_s), Integer
					.parseInt(end_row_s));
			end_column = PublicUtil.max(Integer.parseInt(start_column_s),
					Integer.parseInt(end_column_s));
		} catch (Exception e) {
			System.out.println("解析动态区域的行列设定出错，请确认");
			return;
		}

		// 区域样式，目前只设定边框样式
		Element range_style_info = element.element("style");
		BorderStyle range_borderStyle = new BorderStyle(range_style_info);
		// 设定区域的边框
		table_obj.SetRangeBorderStyle(range_borderStyle, start_row,
				start_column, end_row, end_column);

		// 设定当前区域中的所有单元格的属性
		// 由于对每个单元格而言，是继承父对象的样式，所以边框属性无效
		for (int i = start_row; i <= end_row; i++) {
			for (int j = start_column; j <= end_column; j++) {
				Cell cell_obj = table_obj.getCell(i, j);
				cell_obj.setStyle(range_style_info, true);
			}
		}

		Element provider = element.element("data_source").element("provider");
		String file_url = provider.attributeValue("file_url");
		String class_name = provider.attributeValue("class_name");
		String options = provider.attributeValue("options");
		Element parameters = element.element("data_source").element(
				"parameters");

		ZTableRangeProvider tableProvider = null;

		try {
			Class cls = PublicUtil.loadClass(file_url, class_name);
			tableProvider = (ZTableRangeProvider) cls.newInstance();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return;
		}

		// 传递特殊设定
		tableProvider.set_options(options);

		// 通过回调函数设定动态区域的其他属性
		tableProvider.set_start_row(start_row);
		tableProvider.set_start_column(start_column);
		tableProvider.set_end_row(end_row);
		tableProvider.set_end_column(end_column);

		// 获取并传递参数列表
		if (parameters != null) {
			List<Parameter> para_list = new ArrayList();
			for (Element p : (List<Element>) parameters.elements("paramater")) {
				String name = p.attributeValue("name");
				String value = p.attributeValue("value");
				String remark = p.attributeValue("remark");
				Parameter para = new Parameter(name, value, remark);
				para_list.add(para);
			}
			tableProvider.set_parameters(para_list);
		}

		// 得到返回数据
		TableRange_data tableRange_data = null;
		try {
			tableRange_data = (TableRange_data) tableProvider.get_range_data();
		} catch (Exception e) {
			System.out.println("获取动态表格数据失败，可能是接口不匹配：" + e.getMessage());
		}

		if (tableRange_data == null) {
			return;
		}

		// ----------------------------------------------
		// --- 根据table_data设定当前range中的数据
		// ----------------------------------------------

		// 应该首先设定TableRange的样式，目前暂不支持

		// 首先设定cell
		List<Cell_definition> cell_infos = tableRange_data.get_cell_infos();
		if (cell_infos != null) {
			for (Cell_definition cell_info : cell_infos) {
				// 获取当前cell在表格中的绝对坐标
				// Range中的单元格均为相对坐标，需要转换为绝对坐标
				int absolut_row = start_row + cell_info.get_row() - 1;

				int absolut_column = start_column + cell_info.get_column() - 1;

				List<Data_unit> cell_datas = cell_info.get_cell_datas();

				if (cell_datas != null) {
					// 目标row比当前表格已有的row大，则新建row
					if (absolut_row > table_obj.getRowCount()) {
						table_obj.addRow(absolut_row - table_obj.getRowCount());
					}

					// 目标row比当前表格已有的row大，则新建row
					if (absolut_column > table_obj.getColumnCount()) {
						table_obj.addColumn(absolut_column
								- table_obj.getColumnCount());
					}
					Cell cell_obj = table_obj.getCell(absolut_row,
							absolut_column);

					// 设定整个Cell的样式
					Element cell_style_info = cell_info.get_style_info();
					cell_obj.setStyle(cell_style_info, true);

					// 接着设定单元格中每个数据单元的样式
					for (Data_unit cell_data : cell_datas) {
						Element data_style_info = cell_data.get_style_info();

						if (cell_data.get_type().equalsIgnoreCase("text")) {
							// 获取当前cell数据的样式。一个cell内可能有多个样式
							FontStyle cell_fontStyle = new FontStyle(
									data_style_info);
							ParaStyle cell_paraStyle = new ParaStyle(
									data_style_info);
							_word_instance.appendText((String) cell_data
									.get_data(), data_style_info);
						} else {
							throw new Exception("不支持的cell数据类型: "
									+ cell_data.get_type());
						}
					}
				}
			}
		}

		// 合并Range级别的Cell
		List<Merge_definition> merge_info_list = tableRange_data
				.get_merge_infos();
		if (merge_info_list != null) {
			for (Merge_definition merge_info : merge_info_list) {
				table_obj.mergeCells(merge_info.get_start_row(), merge_info
						.get_start_column(), merge_info.get_end_row(),
						merge_info.get_end_column());
			}
		}
	}


	public String get_file_path_name() {
		return _file_path_name;
	}

	public void set_file_path_name(String _file_path_name) {
		this._file_path_name = _file_path_name;
	}

	public String get_creator() {
		return _creator;
	}

	public void set_creator(String _creator) {
		this._creator = _creator;
	}

	public String get_create_time() {
		return _create_time;
	}

	public void set_create_time(String _create_time) {
		this._create_time = _create_time;
	}

	public String get_last_modifier() {
		return _last_modifier;
	}

	public void set_last_modifier(String _last_modifier) {
		this._last_modifier = _last_modifier;
	}

	public String get_last_modtime() {
		return _last_modtime;
	}

	public void set_last_modtime(String _last_modtime) {
		this._last_modtime = _last_modtime;
	}

	public String get_title() {
		return _title;
	}

	public void set_title(String _title) {
		this._title = _title;
	}

	public String get_remark() {
		return _remark;
	}

	public void set_remark(String _remark) {
		this._remark = _remark;
	}
}
