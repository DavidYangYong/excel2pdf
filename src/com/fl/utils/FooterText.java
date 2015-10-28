package com.fl.utils;

import java.io.Serializable;

import com.itextpdf.text.Element;

/**
 * 定义页脚信息
 * 
 * @author Tendy
 *         2007.7
 */
public class FooterText implements Serializable {
	
	private static final long serialVersionUID = 2503474144970733220L;
	
	/** 中间对齐 */
	public static final int ALIGN_CENTER = Element.ALIGN_CENTER;
	/** 左对齐 */
	public static final int ALIGN_LEFT = Element.ALIGN_LEFT;
	/** 右对齐 */
	public static final int ALIGN_RIGHT = Element.ALIGN_RIGHT;
	
	/** 默认字体大小 */
	public static final float DEFAULT_SIZE = 12.0f;
	
	/** 页码样式：Page 1 of 10 */
	public static final String STYLE_PAGE_NUMBER_N_OFTOTAL = "Page #N of #T";
	/** 页码样式：- 1 - */
	public static final String STYLE_PAGE_NUMBER_N = "- #N -";
	/** 页码样式：第 1 页 */
	public static final String STYLE_PAGE_NUMBER_N_CH = "第 #N 页";
	/** 页码样式：◇ 1 ◇ */
	public static final String STYLE_PAGE_NUMBER_N_CH2 = "◇ #N ◇";
	/** 页码样式：第 1 页，共 10 页 */
	public static final String STYLE_PAGE_NUMBER_N_OFTOTAL_CH = "第 #N 页，共 #T 页";
	/** 代表样式里的 当前页码 */
	public static final String SIGN_PAGE_NUMBER = "#N";
	/** 代表样式里的 总页码 */
	public static final String SIGN_TOTAL_NUMBER = "#T";
	
	/** 文字 */
	private String text;
	
	/** 对齐方式 */
	private int align = ALIGN_LEFT;
	
	/** 是否粗体 */
	private boolean bold;
	
	/** 字体大小 */
	private float fontSize = DEFAULT_SIZE;
	
	public FooterText() {
	}
	
	public FooterText(String text) {
		this.text = text;
	}
	
	public int getAlign() {
		return align;
	}
	
	public void setAlign(int align) {
		this.align = align;
	}
	
	public boolean isBold() {
		return bold;
	}
	
	public void setBold(boolean bold) {
		this.bold = bold;
	}
	
	public float getFontSize() {
		return fontSize;
	}
	
	public void setFontSize(float size) {
		if (size >= 4.0f && size <= 40.0f)
			this.fontSize = size;
	}
	
	public String getText() {
		return text;
	}
	
	public void setText(String text) {
		this.text = text;
	}
	
}
