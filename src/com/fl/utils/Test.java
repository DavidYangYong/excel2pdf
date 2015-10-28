package com.fl.utils;

import com.itextpdf.text.PageSize;

public class Test {
	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception {
		Test test = new Test();
		String rootPath = test.getClass().getResource("/").getPath();
		String source = rootPath + "test.xls";
		String dest = rootPath + "dest.pdf";
		Excel xls = new Excel(source);
		xls.setReportHeaderStartRow(0); // Report header 开始行：第一行
		xls.setReportHeaderEndRow(1); // Report header 结束行：第二行
		xls.setShowPageNumber(false); // 设置显示页码
		xls.setPageSize(PageSize.A4.rotate()); // 设置页大小
		// 可以自定义显示页码
		// xls.setPageNumberStyle("第 " + FooterText.SIGN_PAGE_NUMBER + " 页，共 " +
		// FooterText.SIGN_TOTAL_NUMBER + " 页");
		xls.setPageNumberStyle(FooterText.STYLE_PAGE_NUMBER_N_OFTOTAL_CH);
		xls.setPageNumberStyle("- " + FooterText.SIGN_PAGE_NUMBER + " / "
				+ FooterText.SIGN_TOTAL_NUMBER + " -");
		// xls.setPageNumberFontSize(20);
		xls.setPageHeader(0, 2, 2); // 设置页头，第 1 个 sheet，第三行
		// xls.setPageHeader(1, 0, 1); // 设置页头，第 2 个 sheet，第 1-2 行
		
		Excel2PdfTransformer transformer = new Excel2PdfTransformer(xls);
		transformer.write(dest);
	}
}
