package com.fl.utils;

import java.io.File;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.itextpdf.text.PageSize;
import com.itextpdf.text.Rectangle;

/**
 * 定义 Excel 文件信息
 * 
 * @author Tendy
 *         2007.7
 */
public class Excel implements Serializable {
	
	/**
	 * define start row and end row
	 */
	public class RowRange implements Serializable {
		
		private static final long serialVersionUID = -6329839807511406729L;
		
		private int startRow;
		
		private int endRow;
		
		public RowRange() {
		}
		
		public RowRange(int start, int end) {
			this.startRow = start;
			this.endRow = end;
		}
		
		public int getEndRow() {
			return endRow;
		}
		
		public void setEndRow(int endRow) {
			this.endRow = endRow;
		}
		
		public int getStartRow() {
			return startRow;
		}
		
		public void setStartRow(int startRow) {
			this.startRow = startRow;
		}
	}
	
	private static final long serialVersionUID = -407647832147105795L;
	
	/** 文件 */
	private File file;
	/** 页头设置 */
	private Map pageHeaderSetting = new HashMap();
	
	/** 页脚设置 */
	private List footerTexts = new ArrayList();
	
	/** report header 起始行 */
	private int reportHeaderStartRow = -1;
	/** report header 结束行 */
	private int reportHeaderEndRow = -1;
	/** 是否显示页码 */
	private boolean showPageNumber;
	/** 页码定位 */
	private int pageNumberAlign = FooterText.ALIGN_CENTER;
	/** 页码字体大小 */
	private float pageNumberFontSize = FooterText.DEFAULT_SIZE;
	/** 页码样式 */
	private String pageNumberStyle = FooterText.STYLE_PAGE_NUMBER_N;
	/** 页大小 */
	private Rectangle pageSize = PageSize.A4;
	
	public Excel() {
	}
	
	public Excel(String fileName) {
		this.file = new File(fileName);
	}
	
	public Excel(File file) {
		this.file = file;
	}
	
	// -------------------------------------------- getter/setter
	public File getFile() {
		return file;
	}
	
	public void setFile(File file) {
		this.file = file;
	}
	
	public void setFileName(String fileName) {
		this.file = new File(fileName);
	}
	
	public int getReportHeaderEndRow() {
		return reportHeaderEndRow;
	}
	
	public void setReportHeaderEndRow(int reportHeaderEndRow) {
		if (reportHeaderEndRow < reportHeaderStartRow)
			throw new IllegalArgumentException(
					"reportHeaderEndRow  < reportHeaderStartRow");
		this.reportHeaderEndRow = reportHeaderEndRow;
	}
	
	public int getReportHeaderStartRow() {
		return reportHeaderStartRow;
	}
	
	public void setReportHeaderStartRow(int reportHeaderStartRow) {
		this.reportHeaderStartRow = reportHeaderStartRow;
	}
	
	public boolean isShowPageNumber() {
		return showPageNumber;
	}
	
	public void setShowPageNumber(boolean showPageNumber) {
		this.showPageNumber = showPageNumber;
	}
	
	public int getPageNumberAlign() {
		return pageNumberAlign;
	}
	
	public void setPageNumberAlign(int pageNumberAlign) {
		this.pageNumberAlign = pageNumberAlign;
	}
	
	public float getPageNumberFontSize() {
		return pageNumberFontSize;
	}
	
	public void setPageNumberFontSize(float pageNumberFontSize) {
		this.pageNumberFontSize = pageNumberFontSize;
	}
	
	public String getPageNumberStyle() {
		return pageNumberStyle;
	}
	
	public void setPageNumberStyle(String pageNumberStyle) {
		this.pageNumberStyle = pageNumberStyle;
	}
	
	public Rectangle getPageSize() {
		return pageSize;
	}
	
	public void setPageSize(Rectangle pageSize) {
		if (pageSize != null)
			this.pageSize = pageSize;
	}
	
	// ------------------------------------------- methods
	
	/**
	 * 设置页头
	 * 
	 * @param sheetIndex
	 *            - sheet number, min is 0
	 * @param startRow
	 *            - 开始行（最小是 0）
	 * @param endRow
	 *            - 结束行（最小是0）
	 */
	public void setPageHeader(int sheetIndex, int startRow, int endRow) {
		if (startRow < 0 || endRow < startRow)
			throw new IllegalArgumentException("startRow or endRow is illegal");
		pageHeaderSetting.put(Integer.valueOf(sheetIndex),
				new RowRange(startRow, endRow));
	}
	
	/**
	 * 获得指定 sheet 的 page header 的范围
	 * 
	 * @param sheetIndex
	 *            - sheet number, min is 0
	 * @return
	 */
	public RowRange getPageHeader(int sheetIndex) {
		return (RowRange) pageHeaderSetting.get(Integer.valueOf(sheetIndex));
	}
	
	/**
	 * 设置页头
	 * 
	 * @param sheetIndex
	 *            - sheet number, min is 0
	 * @param range
	 *            - 范围
	 */
	public void setPageHeader(int sheetIndex, RowRange range) {
		if (range == null || range.getEndRow() < 0
				|| range.getEndRow() < range.getStartRow())
			throw new IllegalArgumentException("range");
		pageHeaderSetting.put(Integer.valueOf(sheetIndex), range);
	}
	
	/**
	 * 增加页脚
	 * 
	 * @param text
	 *            - 页脚设置
	 */
	public void addPageFooter(FooterText text) {
		if (text == null || text.getText() == null)
			throw new IllegalArgumentException("text");
		footerTexts.add(text);
	}
	
	/**
	 * 删除页脚
	 * 
	 * @param text
	 *            - 页脚设置
	 */
	public void removePageFooter(FooterText text) {
		if (text != null)
			footerTexts.remove(text);
	}
	
	// 非 public 方法
	List getPageFooter() {
		return footerTexts;
	}
	
	/**
	 * 清除page header 设置
	 */
	public void clearPageHeader() {
		this.pageHeaderSetting.clear();
	}
	
	/**
	 * 清除 report header 设置
	 */
	public void clearReportHeader() {
		reportHeaderStartRow = -1;
		reportHeaderEndRow = -1;
	}
	
	/**
	 * 判断是否设置了 page header
	 * 
	 * @param sheetIndex
	 *            - sheet number
	 * @return
	 */
	public boolean hasPageHeader(int sheetIndex) {
		if (pageHeaderSetting.size() > 0)
			return pageHeaderSetting.get(Integer.valueOf(sheetIndex)) != null;
		return false;
	}
	
	/**
	 * 判断是否设置了 page footer
	 * 
	 * @param sheetIndex
	 *            - sheet number
	 * @return
	 */
	public boolean hasPageFooter(int sheetIndex) {
		return footerTexts.size() > 0;
	}
	
	/**
	 * 判断是否设置了 report header
	 * 
	 * @return
	 */
	public boolean hasReportHeader() {
		return reportHeaderStartRow >= 0
				&& reportHeaderEndRow >= reportHeaderStartRow;
	}
	
}
