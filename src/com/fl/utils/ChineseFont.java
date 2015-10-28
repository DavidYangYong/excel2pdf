package com.fl.utils;

import java.util.regex.Pattern;

import com.itextpdf.text.pdf.BaseFont;

/**
 * 中文字体设置
 * 
 * @author Tendy
 *         2007.7
 */
public final class ChineseFont {
	/** 中文字体 */
	public static BaseFont BASE_CHINESE_FONT = null;
	
	// 初始化中文字体
	static {
		try {
			// 装载中文字体，可能需要1-3秒的时间
			BASE_CHINESE_FONT = BaseFont.createFont("STSong-Light",
					"UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
		} catch (Exception e) {
		}
	}
	
	/**
	 * 判断字符串是否有中文
	 * 
	 * @param s
	 *            - 要判断的字符串
	 * @return
	 */
	public static boolean containsChinese(String s) {
		if (s == null || s.length() == 0)
			return false;
		return Pattern.compile("[\u0391-\uFFE5]+").matcher(s).find();
	}
	
	private ChineseFont() {
	}
}
