package com.fl.utils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import com.fl.utils.Excel.RowRange;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import jxl.Cell;
import jxl.CellType;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.BoldStyle;
import jxl.format.BorderLineStyle;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;

/**
 * Excel 文件转换为 PDF 文件，不支持图表
 * 
 * @author Tendy
 *         2007.7
 */
public class Excel2PdfTransformer {
	/** 页头 table */
	protected PdfPTable headerTable = null;
	
	/** Excel 设置 */
	protected Excel xls = null;
	
	/** 要创建的 PDF 文档 */
	protected Document document = null;
	
	/** 临时文档 */
	protected Document tmpDocument = null;
	
	/** 当前处理的 Excel 工作表 */
	protected Sheet currentSheet = null;
	
	/** 当前 Excel 工作表的合并单元格设置 */
	protected Map currentRegionMap = null;
	
	/** 是否已经生成 Document 的 header */
	protected boolean documentHeaderGenerated = false;
	
	/** 当 Excel 的 border 是 NONE 是，pdf 的 border 是否是 0 */
	protected boolean noEmptyBorder = true;
	
	/** 临时存放 pdfCell */
	protected Map mergeRowCache = new HashMap();
	
	/**
	 * Constructor
	 * 
	 * @param xls
	 *            - Excel 设置
	 */
	public Excel2PdfTransformer(Excel xls) {
		this.xls = xls;
	}
	
	/**
	 * 写到 PDF 文件
	 * 
	 * @param fileName
	 *            - PDF 文件名
	 * @throws Exception
	 */
	public void write(String fileName) throws Exception {
		OutputStream output = new FileOutputStream(fileName);
		try {
			write(output);
		} finally {
			if (output != null) {
				try {
					output.close();
				} catch (Exception e) {
				}
				;
			}
		}
	}
	
	/**
	 * 写到 PDF 文件
	 * 
	 * @param file
	 *            - PDF 文件
	 * @throws Exception
	 */
	public void write(File file) throws Exception {
		OutputStream output = new FileOutputStream(file);
		try {
			write(output);
		} finally {
			if (output != null) {
				try {
					output.close();
				} catch (Exception e) {
				}
				;
			}
		}
	}
	
	/**
	 * 写到流
	 * 
	 * @param output
	 *            - 输出流
	 * @throws Exception
	 */
	public void write(OutputStream output) throws Exception {
		documentHeaderGenerated = false;
		noEmptyBorder = true;
		// mergeRowCache.clear();
		
		InputStream input = new FileInputStream(xls.getFile());
		
		// 读取 Excel 文件
		Workbook workbook = Workbook.getWorkbook(input);
		try {
			if (workbook.getNumberOfSheets() == 0) {
				return;
			}
			
			// 创建 PDF document
			document = new Document(xls.getPageSize(), 50, 50, 50, 50);
			tmpDocument = new Document(PageSize.A4.rotate());
			
			PdfWriter writer = PdfWriter.getInstance(document, output);
			PdfWriter tmpWriter = PdfWriter.getInstance(tmpDocument,
					new ByteArrayOutputStream());
					
			// 新建事件
			CustomPdfPageEvent pageEvent = new CustomPdfPageEvent();
			pageEvent.setWritePageNumber(xls.isShowPageNumber());
			pageEvent.setSkipFirstWrite(true);
			pageEvent.setFooterText(xls.getPageFooter());
			pageEvent.setPageNumberAlign(xls.getPageNumberAlign());
			pageEvent.setPageNumberSize(xls.getPageNumberFontSize());
			pageEvent.setPageNumberStyle(xls.getPageNumberStyle());
			// 设置事件处理
			writer.setPageEvent(pageEvent);
			
			// 打开文档
			document.open();
			tmpDocument.open();
			
			// A document cannot be empty. When closing an empty document,
			// exception will occured.
			Chunk ck = new Chunk("test", new Font());
			tmpDocument.add(ck);
			
			for (int sheetIndex = 0; sheetIndex < workbook
					.getNumberOfSheets(); sheetIndex++) {
				currentSheet = workbook.getSheet(sheetIndex);
				if (currentSheet.getRows() == 0)
					continue;
					
				// 处理合并的单元格
				currentRegionMap = readMergedCells(currentSheet);
				// 清除临时 cache
				if (mergeRowCache.size() > 0)
					mergeRowCache.clear();
					
				// 处理 document header，位于第一个 sheet
				if (sheetIndex == 0) {
					processDocumentHeader();
				} else {
					// start on next page
					pageEvent.setHeader(null);
					pageEvent.setResetMargin(true);
					document.add(Chunk.NEXTPAGE);
				}
				
				// 处理 page header
				headerTable = processPageHeader(sheetIndex);
				// processPageHeader 函数也使用了 currentRegionMap，恢复他们的值
				resetMergeCellInfo(currentRegionMap);
				
				// 把 table 写到临时文档，是为了确定 table 的高度。所有的 table 必须要 render 一次才知道
				// totalHeight
				if (headerTable != null) {
					headerTable.setTotalWidth(
							tmpDocument.right() - tmpDocument.left());
					headerTable.writeSelectedRows(0, -1,
							tmpDocument.leftMargin(), 500,
							tmpWriter.getDirectContent());
				}
				
				// 设置 header
				pageEvent.setHeader(headerTable);
				// 第一页不显示 header
				pageEvent.setSkipFirstWrite(true);
				// 重新设置 margin
				pageEvent.setResetMargin(true);
				
				// 创建表格
				PdfPTable currentTable = new PdfPTable(
						currentSheet.getColumns());
				currentTable.setWidthPercentage(100.0f);
				currentTable.getDefaultCell().setPadding(1.0f);
				currentTable.getDefaultCell().setBorderWidth(0.5f);
				
				int[] currentTableWidths = new int[currentSheet.getColumns()];
				// int[] heights = new int[sheet.getRows()];
				
				for (int i = 0; i < currentSheet.getRows(); i++) {
					// heights[i] = sheet.getRowView(i).getSize();
					if (sheetIndex == 0 && xls.hasReportHeader()
							&& i >= xls.getReportHeaderStartRow()
							&& i <= xls.getReportHeaderEndRow()) {
						// 跳过 document header
						continue;
					}
					processRow(currentTable, i, currentTableWidths);
				}
				
				currentTable.setWidths(currentTableWidths);
				document.add(currentTable);
				
				// 处理图像
				// for (int i = 0; i < currentSheet.getNumberOfImages(); i++) {
				// jxl.Image jxlImage = currentSheet.getDrawing(i);
				// Image iTextImage = Image
				// .getInstance(jxlImage.getImageData());
				// document.add(iTextImage);
				// }
			} // end for (sheetIndex)
			
			// if (pageEvent.isEmptyDocument()) {
			// // 如果没内容，添加内容，否则可能会出错
			// ck = new Chunk("EMPTY DOCUMENT", new Font());
			// document.add(ck);
			// }
			
		} finally {
			// 关闭 workbook
			if (workbook != null)
				workbook.close();
			// 关闭 document
			if (document != null)
				document.close();
			if (tmpDocument != null)
				tmpDocument.close();
			if (input != null) {
				input.close();
			}
		}
	}
	
	/**
	 * 内部类，记录合并单元格的范围及处理状态
	 */
	protected static class RegionInfo {
		public RegionInfo() {
		}
		
		public RegionInfo(Range range) {
			this.range = range;
		}
		
		// 是否已处理
		private boolean done = false;
		
		// 合并单元格的范围
		private Range range;
		
		public boolean isDone() {
			return done;
		}
		
		public Range getRange() {
			return range;
		}
		
		public void setRange(Range range) {
			this.range = range;
		}
		
		public void setDone(boolean done) {
			this.done = done;
		}
	}
	
	// --------------------------------------------------- assistant functions
	
	/**
	 * 处理文档头部，只处理一次
	 * 
	 * @throws Exception
	 */
	protected void processDocumentHeader() throws Exception {
		if (xls.hasReportHeader() && !documentHeaderGenerated) {
			// 默认如果 border 是 empty，不显示 border
			noEmptyBorder = false;
			PdfPTable headerTable = new PdfPTable(currentSheet.getColumns());
			int[] tableWidths = new int[currentSheet.getColumns()];
			
			headerTable.setWidthPercentage(100.0f);
			headerTable.getDefaultCell().setBorderWidth(0.5f);
			
			for (int i = xls.getReportHeaderStartRow(); i <= xls
					.getReportHeaderEndRow(); i++) {
				processRow(headerTable, i, tableWidths);
			}
			
			headerTable.setWidths(tableWidths);
			document.add(headerTable);
			noEmptyBorder = true;
			documentHeaderGenerated = true;
		}
	}
	
	/**
	 * 处理 page header
	 * 
	 * @param sheetIndex
	 *            - 工作表索引，最小是 0
	 * @return
	 * @throws Exception
	 */
	protected PdfPTable processPageHeader(int sheetIndex) throws Exception {
		if (xls.hasPageHeader(sheetIndex)) {
			PdfPTable table = new PdfPTable(currentSheet.getColumns());
			int[] tableWidths = new int[currentSheet.getColumns()];
			table.setWidthPercentage(100.0f);
			table.getDefaultCell().setBorderWidth(0.5f);
			RowRange range = xls.getPageHeader(sheetIndex);
			for (int i = range.getStartRow(); i <= range.getEndRow(); i++) {
				processRow(table, i, tableWidths);
			}
			table.setWidthPercentage(100.0f);
			table.setWidths(tableWidths);
			return table;
		}
		return null;
	}
	
	/**
	 * 处理 Excel 文件的一行
	 * 
	 * @param table
	 *            - PdfPTable
	 * @param i
	 *            - 行号
	 * @param widths
	 *            - 记录每列的宽度
	 */
	protected void processRow(PdfPTable table, int i, int[] widths) {
		int j;
		Cell[] rowCells = currentSheet.getRow(i);
		for (j = 0; j < rowCells.length; j++) {
			if (widths[j] <= 0)
				widths[j] = currentSheet.getColumnView(j).getSize();
				
			PdfPCell pdfCell = null;
			Paragraph content = null;
			Cell cell = rowCells[j];
			jxl.format.CellFormat format = cell.getCellFormat();
			Font font = null;
			if (format != null && format.getFont() != null) {
				font = convertFont(format.getFont());
			} else {
				font = new Font();
			}
			
			String key = i + "," + j;
			// 处理合并单元格
			boolean mergeRow = false;
			if (currentRegionMap.containsKey(key)) {
				RegionInfo info = (RegionInfo) currentRegionMap.get(key);
				Range range = info.getRange();
				Cell topLeft = range.getTopLeft();
				Cell bottomRight = range.getBottomRight();
				if (info.isDone()) {
					// 这个 cell (处于行 i 列 j) 已处理，继续下一个 cell
					if (i > topLeft.getRow()
							&& bottomRight.getRow() - topLeft.getRow() > 0) {
						pdfCell = new PdfPCell();
						// 有 rowspan，从下一行起
						pdfCell.addElement(new Chunk(" "));
						PdfPCell lastRowCell = (PdfPCell) mergeRowCache
								.get(key);
						boolean isLastRow = (i == bottomRight.getRow()); // 是否是合并单元格的最后一行
						
						// 下边框
						if (isLastRow) {
							if (lastRowCell.getBorderWidthTop() > 0.01f) {
								pdfCell.setBorderColorBottom(
										lastRowCell.getBorderColorTop()); // 不是
																			// =
																			// color
																			// bottom，因为
																			// color
																			// bottom
																			// 是
																			// null
								pdfCell.setBorderWidthBottom(0.5f);
							}
						} else {
							pdfCell.setBorderWidthBottom(0.0f);
						}
						
						// 背景色
						pdfCell.setBackgroundColor(
								lastRowCell.getBackgroundColor());
								
						// 设置边框
						pdfCell.setBorderWidthTop(1.0f); // 上边框为 0
						if (topLeft.getColumn() < bottomRight.getColumn()) {
							// 大于 1 列
							if (j == topLeft.getColumn()) {
								// 左边界的 acell
								pdfCell.setBorderWidthRight(0.0f);
								pdfCell.setBorderWidthLeft(
										lastRowCell.getBorderWidthLeft());
								pdfCell.setBorderColorLeft(
										lastRowCell.getBorderColorLeft());
							} else if (j == bottomRight.getColumn()) {
								// 右边界的 cell
								pdfCell.setBorderWidthLeft(0.0f);
								pdfCell.setBorderWidthRight(
										lastRowCell.getBorderWidthRight());
								pdfCell.setBorderColorRight(
										lastRowCell.getBorderColorRight());
							} else {
								// 中间的 cell
								pdfCell.setBorderWidthLeft(0.0f);
								pdfCell.setBorderWidthRight(0.0f);
							}
						} else {
							pdfCell.setBorderWidthLeft(
									lastRowCell.getBorderWidthLeft());
							pdfCell.setBorderColorLeft(
									lastRowCell.getBorderColorLeft());
							pdfCell.setBorderWidthRight(
									lastRowCell.getBorderWidthRight());
							pdfCell.setBorderColorRight(
									lastRowCell.getBorderColorRight());
						}
						
						table.addCell(pdfCell);
					}
					continue;
				} else {
					content = new Paragraph(cell.getContents(), font);
					pdfCell = new PdfPCell(content);
					// 设置单元格合并
					pdfCell.setColspan(
							bottomRight.getColumn() - topLeft.getColumn() + 1);
					if (bottomRight.getRow() > topLeft.getRow()) {
						mergeRow = true;
					}
					
					// 设置单元格状态
					for (int row = topLeft.getRow(); row <= bottomRight
							.getRow(); row++) {
						for (int col = topLeft.getColumn(); col <= bottomRight
								.getColumn(); col++) {
							key = row + "," + col;
							info = (RegionInfo) currentRegionMap.get(key);
							info.setDone(true); // 已处理
							mergeRowCache.put(key, pdfCell); // 把这个 cell
																// 缓存起来，这样可以取它的样式
						}
						
					}
				}
			}
			
			if (cell.getType() == CellType.EMPTY) {
				// 空单元格
				pdfCell = new PdfPCell(new Paragraph(""));
				transferFormat(pdfCell, cell, mergeRow);
				table.addCell(pdfCell);
				// table.addCell(" ");
				continue;
			}
			
			if (pdfCell == null) {
				content = new Paragraph(cell.getContents(), font);
				pdfCell = new PdfPCell(content);
			}
			
			transferFormat(pdfCell, cell, mergeRow);
			
			// pdfCell.setPadding(3.0f);
			// pdfCell.setPaddingBottom(pdfCell.getPaddingTop() + 2.0f);
			table.addCell(pdfCell);
		} // end for (j)
		
		if (j + 1 <= currentSheet.getColumns()) {
			/**
			 * jxl 获得的 Excel 列数并不固定,
			 * sheet.getColumns() 获得最大列数
			 * 每行的列数 sheet.getRow(i) 是 实际列数
			 */
			for (int counter = j + 1; counter <= currentSheet
					.getColumns(); counter++) {
				// 增加一个空白 cell
				table.addCell(" ");
			}
		}
	}
	
	/**
	 * 读取合并单元格
	 * 
	 * @param sheet
	 *            - Excel Sheet
	 */
	protected Map readMergedCells(Sheet sheet) {
		/*
		 * regionMap("行,列") = RegionInfo
		 * 表示 excel 表中，位于 (行,列) 处的单元格，属于被合并的单元格
		 */
		Map regionMap = new HashMap(16);
		Range[] mergedCells = sheet.getMergedCells();
		if (mergedCells != null) {
			for (int i = 0; i < mergedCells.length; i++) {
				Cell topLeft = mergedCells[i].getTopLeft();
				Cell bottomRight = mergedCells[i].getBottomRight();
				for (int row = topLeft.getRow(); row <= bottomRight
						.getRow(); row++) {
					for (int col = topLeft.getColumn(); col <= bottomRight
							.getColumn(); col++) {
						regionMap.put(row + "," + col,
								new RegionInfo(mergedCells[i]));
					}
				}
			}
		}
		return regionMap;
	}
	
	/**
	 * 重新设置 regionMap 为未处理
	 * 
	 * @param regionMap
	 */
	protected void resetMergeCellInfo(Map regionMap) {
		Iterator iter = regionMap.values().iterator();
		while (iter.hasNext()) {
			RegionInfo info = (RegionInfo) iter.next();
			info.setDone(false);
		}
	}
	
	/**
	 * 转换单元格格式
	 * 
	 * @param pdfCell
	 *            - PdfPCell
	 * @param cell
	 *            - jxl.Cell
	 * @param mergeRow
	 *            - 是否合并行
	 */
	protected void transferFormat(PdfPCell pdfCell, Cell cell, boolean mergeRow) {
		jxl.format.CellFormat format = cell.getCellFormat();
		if (format != null) {
			// 水平对齐
			pdfCell.setHorizontalAlignment(
					convertAlignment(format.getAlignment(), cell.getType()));
			// 垂直对齐
			pdfCell.setVerticalAlignment(
					convertVerticalAlignment(format.getVerticalAlignment()));
					// 背景
					// if (format.getBackgroundColour() != null) {
					// pdfCell.setBackgroundColor(convertColour(
					// format.getBackgroundColour(), Color.WHITE));
					// }
					
			// 处理 border
			BorderLineStyle lineStyle = null;
			if (mergeRow) {
				pdfCell.setBorderWidthBottom(0.0f);
			} else {
				lineStyle = format.getBorderLine(jxl.format.Border.BOTTOM);
				pdfCell.setBorderWidthBottom(convertBorderStyle(lineStyle));
				if (lineStyle.getValue() == BorderLineStyle.NONE.getValue())
					pdfCell.setBorderColorBottom(BaseColor.GRAY);
				// else
				// pdfCell.setBorderColorBottom(convertColour(
				// format.getBorderColour(jxl.format.Border.BOTTOM),
				// Color.GRAY));
			}
			
			lineStyle = format.getBorderLine(jxl.format.Border.TOP);
			pdfCell.setBorderWidthTop(convertBorderStyle(lineStyle));
			// if (lineStyle.getValue() == BorderLineStyle.NONE.getValue())
			// pdfCell.setBorderColorTop(Color.GRAY);
			// else
			// pdfCell.setBorderColorTop(convertColour(
			// format.getBorderColour(jxl.format.Border.TOP),
			// Color.GRAY));
			
			lineStyle = format.getBorderLine(jxl.format.Border.LEFT);
			pdfCell.setBorderWidthLeft(convertBorderStyle(lineStyle));
			// if (lineStyle.getValue() == BorderLineStyle.NONE.getValue())
			// pdfCell.setBorderColorLeft(Color.GRAY);
			// else
			// pdfCell.setBorderColorLeft(convertColour(
			// format.getBorderColour(jxl.format.Border.LEFT),
			// Color.GRAY));
			
			lineStyle = format.getBorderLine(jxl.format.Border.RIGHT);
			pdfCell.setBorderWidthRight(convertBorderStyle(lineStyle));
			// if (lineStyle.getValue() == BorderLineStyle.NONE.getValue())
			// pdfCell.setBorderColorRight(Color.GRAY);
			// else
			// pdfCell.setBorderColorRight(convertColour(
			// format.getBorderColour(jxl.format.Border.RIGHT),
			// Color.GRAY));
			
		}
		
	}
	
	/**
	 * 未用的函数
	 * 把数值转换成坐标
	 * 
	 * @param d
	 *            - 数值，jExcel 的 Image.getRow 或 Image.getColumn
	 * @param data
	 *            - jExcel 表格 列长度 或 行高度 的数组
	 * @return
	 */
	protected float getPosition(double d, int[] data) {
		float f = 0.0f;
		int end = (int) d;
		if (end < 0 || end >= data.length)
			return 0.0f;
			
		for (int i = 0; i <= end; i++) {
			f += data[i];
		}
		
		f += (float) (d - end) * data[end];
		return f;
	}
	
	/**
	 * 转换水平对齐
	 * 
	 * @param align
	 *            - jxl 中的对齐方式
	 * @param cellType
	 *            - 单元格类型
	 * @return
	 */
	protected int convertAlignment(Alignment align, CellType cellType) {
		if (align == null)
			return Element.ALIGN_UNDEFINED;
			
		if (Alignment.CENTRE.getValue() == align.getValue())
			return Element.ALIGN_CENTER;
			
		if (Alignment.LEFT.getValue() == align.getValue())
			return Element.ALIGN_LEFT;
			
		if (Alignment.RIGHT.getValue() == align.getValue())
			return Element.ALIGN_RIGHT;
			
		if (Alignment.JUSTIFY.getValue() == align.getValue())
			return Element.ALIGN_JUSTIFIED;
			
		if (Alignment.GENERAL.getValue() == align.getValue()) {
			// 所有未明确设置对齐方式的元素，都属于 Alignment.GENERAL 类型
			if (cellType == CellType.NUMBER
					|| cellType == CellType.NUMBER_FORMULA)
				return Element.ALIGN_RIGHT; // 数字右对齐
			if (cellType == CellType.DATE || cellType == CellType.DATE_FORMULA)
				return Element.ALIGN_RIGHT; // 日期右对齐
		}
		return Element.ALIGN_UNDEFINED;
	}
	
	/**
	 * 转换垂直对齐方式
	 * 
	 * @param align
	 *            - jxl 的对齐方式
	 * @return
	 */
	protected int convertVerticalAlignment(VerticalAlignment align) {
		if (align == null)
			return Element.ALIGN_UNDEFINED;
			
		if (VerticalAlignment.BOTTOM.getValue() == align.getValue())
			return Element.ALIGN_BOTTOM;
			
		if (VerticalAlignment.CENTRE.getValue() == align.getValue())
			return Element.ALIGN_MIDDLE;
			
		if (VerticalAlignment.TOP.getValue() == align.getValue())
			return Element.ALIGN_TOP;
			
		if (VerticalAlignment.JUSTIFY.getValue() == align.getValue())
			return Element.ALIGN_JUSTIFIED;
			
		return Element.ALIGN_UNDEFINED;
	}
	
	/**
	 * 转换颜色
	 * 
	 * @param c
	 *            - 要转换的颜色
	 * @param defaultColor
	 *            - 默认颜色，当参数 c 为 null 时，使用默认颜色
	 * @return
	 */
	protected BaseColor convertColour(BaseColor c, BaseColor defaultColor) {
		if (defaultColor == null)
			defaultColor = BaseColor.WHITE;
			
		if (c == null)
			return defaultColor;
			
		return null;
		
		// if (c == BaseColor.AUTOMATIC) // Excel中的自动(前景色)
		// return BaseColor.BLACK;
		// else if (c == BaseColor.) // Excel中的自动(底色)
		// return BaseColor.WHITE;
		
		// RGB rgb = c.getDefaultRGB();
		// return new Color(rgb.getRed(), rgb.getGreen(), rgb.getBlue());
		// return Colour.DEFAULT_BACKGROUND;
	}
	
	/**
	 * 转换字体
	 * 
	 * @param f
	 *            - 字体
	 * @return
	 */
	protected Font convertFont(jxl.format.Font f) {
		if (f == null || f.getName() == null)
			return FontFactory.getFont(FontFactory.COURIER, BaseFont.IDENTITY_H,
					BaseFont.NOT_EMBEDDED);
					
		int fontStyle = convertFontStyle(f);
		Font font = null;
		// Color fontColor = convertColour(f.getColour(), Color.BLACK);
		if (ChineseFont.BASE_CHINESE_FONT != null
				&& ChineseFont.containsChinese(f.getName())) {
			font = new Font(ChineseFont.BASE_CHINESE_FONT);
		} else {
			String s = f.getName().toLowerCase();
			int fontFamily;
			// if (s.indexOf("courier") >= 0) // "courier new".equals(s) ||
			// // "courier".equals(s))
			// // fontFamily = Font.COURIER;
			// else if (s.indexOf("times") >= 0)
			// // fontFamily = Font.TIMES_ROMAN;
			// else
			// // fontFamily = Font.HELVETICA;
			
			font = new Font();
			
		}
		
		return font;
	}
	
	/**
	 * 转换字体样式
	 * 
	 * @param font
	 *            - 字体
	 * @return
	 */
	protected int convertFontStyle(jxl.format.Font font) {
		
		int result = Font.NORMAL;
		if (font.isItalic())
			result |= Font.ITALIC;
			
		if (font.isStruckout())
			result |= Font.STRIKETHRU;
			
		if (font.getBoldWeight() == BoldStyle.BOLD.getValue())
			result |= Font.BOLD;
			
		if (font.getUnderlineStyle() != null) {
			// 下划线
			UnderlineStyle style = font.getUnderlineStyle();
			if (style.getValue() != UnderlineStyle.NO_UNDERLINE.getValue())
				result |= Font.UNDERLINE;
		}
		return result;
	}
	
	/**
	 * 转换边框样式
	 * 
	 * @param style
	 *            - jxl.format.BorderLineStyle
	 * @return
	 */
	protected float convertBorderStyle(BorderLineStyle style) {
		if (style == null)
			return 0.0f;
			
		float w = 0.0f;
		if (BorderLineStyle.NONE.getValue() == style.getValue()) {
			// 默认全部使用边框，边框大小 0.5f
			if (noEmptyBorder)
				w = 0.1f;
		} else if (BorderLineStyle.THIN.getValue() == style.getValue())
			w = 0.1f;
		else if (BorderLineStyle.THICK.getValue() == style.getValue()) {
			w = 1.5f;
		} else if (BorderLineStyle.MEDIUM.getValue() == style.getValue()) {
			w = 1.0f;
		} else {
			w = 0.1f;
		}
		return w;
	}
	
	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception {
		String source = "e:\\新建 Microsoft Excel 工作表.xls";
		String dest = "e:\\b.pdf";
		Excel xls = new Excel(source);
		xls.setReportHeaderStartRow(0); // Report header 开始行：第一行
		xls.setReportHeaderEndRow(1); // Report header 结束行：第二行
		xls.setShowPageNumber(false); // 设置显示页码
		xls.setPageSize(PageSize.A4.rotate()); // 设置页大小
		// 可以自定义显示页码
		// xls.setPageNumberStyle("第 " + FooterText.SIGN_PAGE_NUMBER + " 页，共 " +
		// FooterText.SIGN_TOTAL_NUMBER + " 页");
		xls.setPageNumberStyle(FooterText.STYLE_PAGE_NUMBER_N_OFTOTAL_CH);
		// xls.setPageNumberStyle("- " + FooterText.SIGN_PAGE_NUMBER + " / " +
		// FooterText.SIGN_TOTAL_NUMBER + " -");
		// xls.setPageNumberFontSize(20);
		// xls.setPageHeader(0, 2, 2); // 设置页头，第 1 个 sheet，第三行
		// xls.setPageHeader(1, 0, 1); // 设置页头，第 2 个 sheet，第 1-2 行
		
		Excel2PdfTransformer transformer = new Excel2PdfTransformer(xls);
		transformer.write(dest);
	}
	
}
