package com.zeroapp.zerodoc.msword;

import java.io.File;

import org.dom4j.Element;

import com.jacob2.activeX.ActiveXComponent;
import com.jacob2.com.Dispatch;
import com.jacob2.com.Variant;
import com.zeroapp.zerodoc.exceptions.ZerodocException;
import com.zeroapp.zerodoc.msword.util.BorderStyle;
import com.zeroapp.zerodoc.msword.util.ConstantUtil;

//	get 获取属性
//	put 设定属性
//	invoke 调用sub
//	call 调用function
//	word中很多对象的尺寸是以磅为单位的，显示的是打印出来之后的尺寸，1磅 = 1/72 英寸
//	对象后面加括号的通常是引用对象的Item属性，如Borders.(wdBorderVertical) = Borders.Item(wdBorderVertical)

public class WordInstance {

	// doc对象
	private Dispatch _doc = null;

	@SuppressWarnings("unused")
	private WordInstance() {
	}

	public WordInstance(Dispatch doc) {
		assert (doc != null);
		_doc = doc;
	}

	/**
	 * 保存Word文档到指定的目录(包括文件名)
	 * 
	 * @param filePath
	 */
	public void saveas(final String filePath) {
		// 保存文件
		Dispatch.invoke(_doc, "SaveAs", Dispatch.Method, new Object[] {
				filePath, new Variant(0) }, new int[1]);
	}

	/**
	 * 关闭文档
	 * 
	 * @param flag,
	 *            表示是否保存文档 0 = wdDoNotSaveChanges -1 = wdSaveChanges -2 =
	 *            wdPromptToSaveChanges
	 * 
	 */
	public void close(int flag) {
		// Close the document without saving changes

		Dispatch.call(_doc, "Close", new Variant(flag));

		_doc = null;
	}

	/**
	 * 在文档当前选中区域的末尾追加文本
	 * 
	 * @throws Exception
	 */
	public void appendText(String text, Element style_info) throws Exception {
		try {
			Dispatch selection = Dispatch.get(WordFactory.MsWordApp,
					"Selection").toDispatch();

			Dispatch range = Dispatch.get(selection, "Range").toDispatch();

			Dispatch.call(range, "InsertAfter", text);
			Range range_obj = new Range(range);
			range_obj.setStyle(style_info);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 格式化文档中选中的区域
	 * 
	 * @param style_info
	 * @throws Exception
	 */
	public void formatSelection(Element style_info) throws Exception {
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch range = Dispatch.get(selection, "Range").toDispatch();
		Range range_obj = new Range(range);
		range_obj.setStyle(style_info);

	}

	/**
	 * 在文档当前Range的末尾追加文本
	 * 
	 * @param text
	 * @param style_info
	 * @throws Exception
	 */
	public void insertText(String text, Element style_info) throws Exception {
		Range range = this.getCurrentRange();
		range.appendText(text, style_info);
	}

	/**
	 * 在文档中插入图片
	 * 
	 * @param pic_file_path
	 * @throws ZerodocException
	 */
	public Image addImage(String pic_file_path) throws ZerodocException {
		// Get the current selection within Word at the moment. If
		// a new document has just been created then this will be at
		// the top of the new doc
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		File f = new File(pic_file_path);
		if (!f.exists()) {
			throw new ZerodocException("图片文件" + pic_file_path + "没有找到");
		}

		// 插入图片
		Variant image = Dispatch.call(Dispatch.get(selection, "InLineShapes")
				.toDispatch(), "AddPicture", pic_file_path);

		return new Image(image);
	}

	/**
	 * 打印文档
	 */
	public void printFile() {
		// Just print the current document to the default printer
		Dispatch.call(_doc, "PrintOut");
	}

	/**
	 * 查找替换文本
	 * 
	 * @param searchText
	 *            需要查找的关键字
	 * @param replaceText
	 *            要替换成的关键字
	 */
	public void searchReplace(final String searchText, final String replaceText) {
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		Dispatch find = ActiveXComponent.call(selection, "Find").toDispatch();

		// 查找什么文本
		Dispatch.put(find, "Text", searchText);

		// 替换文本
		Dispatch.call(find, "ClearFormatting");
		Dispatch.put(find, "Text", searchText);
		Dispatch.call(find, "Execute");
		Dispatch.put(selection, "Text", replaceText);
	}

	/**
	 * 把指定的值设置到指定的标签中去
	 * 
	 * @param bookMarkKey
	 */
	public void setBookMark(final String bookMarkKey, final String bookMarkValue) {
		Dispatch bookMarks = ActiveXComponent.call(_doc, "Bookmarks")
				.toDispatch();

		boolean bookMarkExist1 = Dispatch
				.call(bookMarks, "Exists", bookMarkKey).getBoolean();

		if (bookMarkExist1 == true) {
			Dispatch rangeItem = Dispatch.call(bookMarks, "Item", bookMarkKey)
					.toDispatch();
			Dispatch range = Dispatch.call(rangeItem, "Range").toDispatch();

			if (bookMarkValue != null) {
				Dispatch.put(range, "Text", new Variant(bookMarkValue));
			}
		} else {
			System.out.println("not exists bookmark!");
		}
	}

	/**
	 * 添加新一行
	 */
	public void newLine() {
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// Put the specified text at the insertion point
		Dispatch.call(selection, "TypeParagraph");
	}

	/**
	 * 新一页（在末尾加入分页符）
	 */
	public void newPage() {
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// Put the specified text at the insertion point
		Dispatch.call(selection, "InsertBreak", new Variant(
				ConstantUtil.wdPageBreak));
	}

	/**
	 * 新建一个section，可以是下一页或连续
	 * 
	 * @throws Exception
	 */
	public Section addSection(String sectionType) throws Exception {

		// 光标移动到文档最后
		this.moveCursorToEnd();
		
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		if ("wdSectionBreakNextPage".equalsIgnoreCase(sectionType)) {
			// Put the specified text at the insertion point
			Dispatch.call(selection, "InsertBreak", new Variant(
					ConstantUtil.wdSectionBreakNextPage));
		} else {
			// Put the specified text at the insertion point
			Dispatch.call(selection, "InsertBreak", new Variant(
					ConstantUtil.wdSectionBreakContinuous));
		}

		return this.getLastSection();
	}

	/**
	 * 清除文档选中部分的格式设置
	 */
	public void clearFormatting() {
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// 清除格式
		Dispatch.call(selection, "ClearFormatting");

	}

	/**
	 * 为word添加目录
	 */
	public void addContent() {
		Dispatch contents = ActiveXComponent.call(_doc, "TablesOfContents")
				.toDispatch();
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch range = ActiveXComponent.call(selection, "Range").toDispatch();

		Dispatch.call(contents, "Add", range);
	}

	/**
	 * 更新目录
	 */
	public void updateContent() {
		Dispatch activeDoc = Dispatch.get(WordFactory.MsWordApp,
				"ActiveDocument").toDispatch();
		Dispatch fields = ActiveXComponent.call(activeDoc, "Fields")
				.toDispatch();

		Dispatch.call(fields, "Update");
	}

	/**
	 * 移动到文档的开头
	 */
	public void moveCursorToBegin() {
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// 移到文档开头
		Dispatch.call(selection, "HomeKey", new Variant(6), new Variant(0));
	}

	/**
	 * 移动到文档的末尾
	 */
	public void moveCursorToEnd() {
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		// 移到文档末尾
		Dispatch.call(selection, "EndKey", new Variant(6), new Variant(0));
	}

	/**
	 * 设定当前selection的对齐方式
	 * 
	 * @param alignment
	 */
	public void setAlignment(int alignment) {
		// 得到selection对象
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch paragraphFormat = Dispatch.get(selection, "ParagraphFormat")
				.toDispatch();

		Dispatch.put(paragraphFormat, "Alignment", new Variant(alignment));
	}

	/**
	 * 创建表格
	 * 
	 * @param rows
	 * @param columns
	 */
	public Table addTable(int rows, int columns) {
		Dispatch activeDoc = Dispatch.get(WordFactory.MsWordApp,
				"ActiveDocument").toDispatch();
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch tables = ActiveXComponent.call(activeDoc, "Tables")
				.toDispatch();
		Dispatch range = ActiveXComponent.call(selection, "Range").toDispatch();
		Variant table = Dispatch.call(tables, "Add", range, new Variant(rows),
				new Variant(columns));

		Table table_obj = new Table(table);

		// 设定table的缺省边框
		BorderStyle borderStyle = new BorderStyle();
		table_obj.setBorderStyle(borderStyle);

		return table_obj;
	}

	/**
	 * 移动到页眉
	 */
	public void seekToPageHeader() {
		Dispatch activeDoc = Dispatch.get(WordFactory.MsWordApp,
				"ActiveDocument").toDispatch();
		Dispatch activeWindow = ActiveXComponent
				.call(activeDoc, "ActiveWindow").toDispatch();
		Dispatch view = ActiveXComponent.call(activeWindow, "View")
				.toDispatch();

		Dispatch.put(view, "SeekView", new Variant(
				ConstantUtil.wdSeekCurrentPageHeader));
	}

	/**
	 * 移动到页脚
	 */
	public void seekToPageFooter() {
		Dispatch activeDoc = Dispatch.get(WordFactory.MsWordApp,
				"ActiveDocument").toDispatch();
		Dispatch activeWindow = ActiveXComponent
				.call(activeDoc, "ActiveWindow").toDispatch();
		Dispatch view = ActiveXComponent.call(activeWindow, "View")
				.toDispatch();

		Dispatch.put(view, "SeekView", new Variant(
				ConstantUtil.wdSeekCurrentPageFooter));
	}

	/**
	 * 移动到主文档
	 */
	public void seekToMainDocument() {
		Dispatch activeDoc = Dispatch.get(WordFactory.MsWordApp,
				"ActiveDocument").toDispatch();
		Dispatch activeWindow = ActiveXComponent
				.call(activeDoc, "ActiveWindow").toDispatch();
		Dispatch view = ActiveXComponent.call(activeWindow, "View")
				.toDispatch();

		Dispatch.put(view, "SeekView", new Variant(
				ConstantUtil.wdSeekMainDocument));
	}

	/**
	 * 加入页码
	 */
	public void addPageNumber() {
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch fields = ActiveXComponent.call(selection, "Fields")
				.toDispatch();
		Dispatch range = ActiveXComponent.call(selection, "Range").toDispatch();

		Dispatch.callSub(fields, "Add", range, new Variant(
				ConstantUtil.wdFieldPage));
	}

	/**
	 * 加入当前的总页数，允许用户指定偏移量
	 * 
	 * @return
	 * @throws ZerodocException
	 */
	public void addCurrentPageCount(int offset) throws ZerodocException {
		Dispatch activeDoc = Dispatch.get(WordFactory.MsWordApp,
				"ActiveDocument").toDispatch();
		Variant numPages = Dispatch.call(activeDoc,
				"BuiltInDocumentProperties", new Variant(
						ConstantUtil.wdPropertyPages));
		Integer pageCount = numPages.toInt();
		pageCount += offset;
		try {
			Range range = this.getLastRange();
			range.appendText(pageCount.toString(), null);
		} catch (Exception e) {
			throw new ZerodocException("添加自当前总页数失败：" + e.getMessage());
		}
	}

	/**
	 * 加入文档的总页数，不能修改
	 */
	public void addPageCount() {
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch fields = ActiveXComponent.call(selection, "Fields")
				.toDispatch();
		Dispatch range = ActiveXComponent.call(selection, "Range").toDispatch();

		Dispatch.callSub(fields, "Add", range, new Variant(
				ConstantUtil.wdFieldNumPages));
	}

	/**
	 * 加入日期时间
	 */
	public void addDataTime_stamp(String dateTimeFormat, Boolean insertAsField) {

		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();

		String format = dateTimeFormat != null ? dateTimeFormat : "yyyy-M-d";
		boolean asField = insertAsField != null ? insertAsField : true;

		Dispatch.callSub(selection, "InsertDateTime", format, asField);
	}

	/**
	 * 返回最后一个section
	 * 
	 * @return
	 * @throws Exception
	 */
	public Section getLastSection() throws Exception {
		this.moveCursorToEnd();
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Variant section = ActiveXComponent.call(selection, "Sections",
				new Variant(1));

		if (section == null) {
			throw new Exception("unable to find any section.");
		}
		return new Section(section);
	}

	/**
	 * 返回文档末尾的Range
	 * 
	 * @return
	 */
	public Range getLastRange() {
		this.moveCursorToEnd();
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch range = Dispatch.get(selection, "Range").toDispatch();
		int endKey = Dispatch.get(range, "End").getInt();
		Dispatch.call(range, "SetRange", endKey - 1, endKey - 1);
		return new Range(range);
	}

	public Range getCurrentRange() {
		Dispatch selection = Dispatch.get(WordFactory.MsWordApp, "Selection")
				.toDispatch();
		Dispatch range = Dispatch.get(selection, "Range").toDispatch();
		return new Range(range);

	}
}
