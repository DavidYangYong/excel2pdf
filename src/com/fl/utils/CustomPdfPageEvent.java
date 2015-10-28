package com.fl.utils;

import java.util.Iterator;
import java.util.List;

import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * 自定义 PDF 事件
 * 
 * @author Tendy
 *         2007.7
 */
public class CustomPdfPageEvent extends PdfPageEventHelper {
	/** page header */
	private PdfPTable headerTable = null;
	
	/** the footer setting */
	private List footerTexts = null;
	/** 是否有 footer */
	private boolean hasFooter = false;
	/** footer 高度 */
	private float footerHeight = 0.0f;
	
	/** A template that will hold the total number of pages. */
	private PdfTemplate pageNumberTpl = null;
	/** 页脚所用字体 */
	private BaseFont font = null;
	
	/** 是否显示页码 */
	private boolean writePageNumber = false;
	/** 页面字体大小 */
	private float pageNumberSize = 0.0f;
	/** 页码定位 */
	private int pageNumberAlign = Element.ALIGN_CENTER;
	/** 页码样式 */
	private String pageNumberStyle = FooterText.STYLE_PAGE_NUMBER_N;
	/** 是否显示总页码 */
	private boolean showTotalPage = false;
	/** 代替总页数的空格定义，在 openDocument 时计算 */
	private String blankTextChars = "    ";
	
	/** 保存初始的 top margin 值 */
	private float _topMargin = 0.0f;
	/** 保存初始的 bottom margin 值 */
	private float _bottomMargin = 0.0f;
	/** 跳过第一个页头 */
	private boolean skipFirstWrite = false;
	/** 是否重设 margin */
	private boolean resetMargin = false;
	/** 是否是空文档 */
	private boolean emptyDocument = true;
	
	CustomPdfPageEvent() {
	}
	
	/**
	 * 设置是否输出页码
	 * 
	 * @param writePageNumber
	 */
	void setWritePageNumber(boolean writePageNumber) {
		this.writePageNumber = writePageNumber;
	}
	
	/**
	 * 设置页码位置
	 * 
	 * @param align
	 *            - Element.ALIGN_LEFT, Element.ALIGN_CENTER,
	 *            Element.ALIGN_RIGHT
	 */
	void setPageNumberAlign(int align) {
		pageNumberAlign = align;
	}
	
	/**
	 * 设置页码字体大小
	 * 
	 * @param size
	 */
	void setPageNumberSize(float size) {
		pageNumberSize = size;
		if (size > footerHeight)
			footerHeight = size;
	}
	
	/**
	 * 清空 page header
	 */
	void clearHeader() {
		headerTable = null;
	}
	
	/**
	 * 设置是否跳过第一个页头
	 * 
	 * @param skip
	 */
	void setSkipFirstWrite(boolean skip) {
		this.skipFirstWrite = skip;
	}
	
	/**
	 * 返回文档是否为空
	 * 
	 * @return
	 */
	public boolean isEmptyDocument() {
		// 如果 page end 事件没发生过，为空文档
		return emptyDocument;
	}
	
	/**
	 * 设置页头
	 * 
	 * @param table
	 *            - 页头table
	 */
	void setHeader(PdfPTable table) {
		headerTable = table;
	}
	
	/**
	 * 设置页脚
	 * 
	 * @param footerTexts
	 */
	void setFooterText(List footerTexts) {
		this.footerTexts = footerTexts;
		if (footerTexts != null && footerTexts.size() > 0) {
			hasFooter = true;
			Iterator iter = footerTexts.iterator();
			// 找出最大的字体size
			while (iter.hasNext()) {
				FooterText text = (FooterText) iter.next();
				if (text.getFontSize() > footerHeight)
					footerHeight = text.getFontSize();
			}
			if (footerHeight < pageNumberSize)
				footerHeight = pageNumberSize;
		} else {
			hasFooter = false;
			footerHeight = 0.0f;
		}
	}
	
	/**
	 * 设置页码显示的样式
	 * 
	 * @param style
	 */
	void setPageNumberStyle(String style) {
		if (style != null & style.indexOf(FooterText.SIGN_PAGE_NUMBER) >= 0) {
			pageNumberStyle = style;
			showTotalPage = (style.indexOf(FooterText.SIGN_TOTAL_NUMBER) > 0);
			// writePageNumber = true;
		}
	}
	
	/**
	 * 是否重置 margin
	 * 
	 * @param reset
	 */
	void setResetMargin(boolean reset) {
		resetMargin = reset;
	}
	
	/**
	 * 设置 margin
	 * 
	 * @param document
	 *            - com.lowagie.text.Document
	 */
	private void setMargin(Document document) {
		float leftMargin = document.leftMargin();
		float rightMargin = document.rightMargin();
		float topMargin = (headerTable == null) ? this._topMargin
				: this._topMargin + headerTable.getTotalHeight();
		float bottomMargin = this._bottomMargin + footerHeight;
		
		document.setMargins(leftMargin, rightMargin, topMargin, bottomMargin);
	}
	
	// ------------------------------------------ event implementation
	
	/**
	 * 页结束事件
	 */
	public void onEndPage(PdfWriter writer, Document document) {
		Rectangle page = document.getPageSize();
		PdfContentByte cb = writer.getDirectContent();
		if (writePageNumber && pageNumberTpl != null) {
			cb.saveState();
			// compose the footer
			String text = this.pageNumberStyle.replaceAll(
					FooterText.SIGN_PAGE_NUMBER,
					String.valueOf(writer.getPageNumber()));
			int totalPagePos = -1; // 总页码在text中的位置
			if (showTotalPage == true) {
				totalPagePos = text.indexOf(FooterText.SIGN_TOTAL_NUMBER);
				text = text.replaceAll(FooterText.SIGN_TOTAL_NUMBER,
						blankTextChars);
			}
			// 文字占的宽度
			float textSize = font.getWidthPoint(text, pageNumberSize);
			// Y 坐标
			float textBase = document.bottomMargin() - footerHeight;
			cb.beginText();
			cb.setFontAndSize(font, pageNumberSize);
			
			// 计算 X 坐标
			float x = 0.0f;
			if (this.pageNumberAlign == Element.ALIGN_CENTER)
				x = (page.getWidth() - textSize) / 2;
			else if (this.pageNumberAlign == Element.ALIGN_LEFT)
				x = document.left();
			else
				x = document.right() - textSize
						- font.getWidthPoint("00", pageNumberSize);
			cb.setTextMatrix(x, textBase);
			cb.showText(text);
			cb.endText();
			if (showTotalPage == true) {
				textSize = font.getWidthPoint(text.substring(0, totalPagePos),
						pageNumberSize);
				cb.addTemplate(pageNumberTpl, x + textSize, textBase);
			}
			cb.restoreState();
		}
		
		if (hasFooter) {
			// 显示 page footer
			cb.saveState();
			float x = 0.0f;
			float textBase = document.bottomMargin() - footerHeight;
			for (int i = 0; i < footerTexts.size(); i++) {
				FooterText text = (FooterText) footerTexts.get(i);
				
				cb.beginText();
				cb.setFontAndSize(font, text.getFontSize());
				if (text.isBold())
					cb.setTextRenderingMode(
							PdfContentByte.TEXT_RENDER_MODE_FILL_STROKE);
				else
					cb.setTextRenderingMode(
							PdfContentByte.TEXT_RENDER_MODE_FILL);
				if (text.getAlign() == Element.ALIGN_CENTER) {
					x = document.getPageSize().getWidth() / 2;
				} else if (text.getAlign() == Element.ALIGN_LEFT) {
					x = document.left();
				} else {
					x = document.right();
				}
				cb.showTextAligned(text.getAlign(), text.getText(), x, textBase,
						0.0f);
				cb.endText();
			}
			
			cb.restoreState();
		}
		// reset skipFirstWrite
		skipFirstWrite = false;
		
		if (resetMargin) {
			// 设置 margin
			setMargin(document);
			resetMargin = false;
		}
		// not empty document
		emptyDocument = false;
	}
	
	/**
	 * 打开文档事件
	 */
	public void onOpenDocument(PdfWriter writer, Document document) {
		try {
			if (writePageNumber) {
				// initialization of the template
				pageNumberTpl = writer.getDirectContent().createTemplate(100,
						100);
				pageNumberTpl.setBoundingBox(new Rectangle(-20, -20, 100, 100));
				// initialization of the font
				if (ChineseFont.containsChinese(pageNumberStyle)
						&& ChineseFont.BASE_CHINESE_FONT != null)
					font = ChineseFont.BASE_CHINESE_FONT;
				else
					font = BaseFont.createFont(BaseFont.HELVETICA,
							BaseFont.WINANSI, false);
							
				// 计算需要多少空格来代替 "总页数" 的位置
				float size = font.getWidthPoint("000", this.pageNumberSize);
				float blankUnitSize = font.getWidthPoint(" ",
						this.pageNumberSize);
				int needSpaceChars = Math.round(size / blankUnitSize);
				blankTextChars = "";
				for (int i = 0; i < needSpaceChars; i++)
					blankTextChars += " ";
			}
			// 保存初始的 top、bottom margin
			_topMargin = document.topMargin();
			_bottomMargin = document.bottomMargin();
		} catch (Exception e) {
			throw new ExceptionConverter(e);
		}
	}
	
	/**
	 * Start page 事件
	 */
	public void onStartPage(PdfWriter writer, Document document) {
		if (headerTable != null && skipFirstWrite == false) {
			Rectangle page = document.getPageSize();
			headerTable.setTotalWidth(document.right() - document.left());
			headerTable.writeSelectedRows(0, -1, document.leftMargin(),
					page.getHeight() - document.topMargin()
							+ headerTable.getTotalHeight(),
					writer.getDirectContent());
		}
	}
	
	/**
	 * Close document 事件
	 */
	public void onCloseDocument(PdfWriter writer, Document document) {
		if (showTotalPage == true && writePageNumber && pageNumberTpl != null) {
			pageNumberTpl.beginText();
			pageNumberTpl.setFontAndSize(font, pageNumberSize);
			// 调整位置 (x 坐标)
			float x = 0.0f;
			int totalPage = writer.getPageNumber() - 1;
			if (totalPage < 10) // 1 位数
				x += font.getWidthPoint("00", pageNumberSize) / 2;
			else if (totalPage < 100) // 2 位数
				x += font.getWidthPoint("0", pageNumberSize) / 2;
			else if (totalPage > 1000) // 4 位数或更多
				x -= font.getWidthPoint("0", pageNumberSize) / 2;
				
			pageNumberTpl.setTextMatrix(x, 0);
			pageNumberTpl.showText("" + totalPage);
			pageNumberTpl.endText();
		}
	}
}
