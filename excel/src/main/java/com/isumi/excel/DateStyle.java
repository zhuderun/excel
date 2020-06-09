package com.isumi.excel;


/**
 * 日期类型
 * @author zdr
 *
 */
public enum DateStyle {
	YYYY("yyyy"),MM("MM"),DD("dd"),YYYYMM("yyyyMM"),YYYY_MM("yyyy-MM"),MMDD("MMdd"),MM_DD("MM-dd"),YYYYMMDD("yyyyMMdd"),YYYY_MM_DD("yyyy-MM-dd");
	
	private String style;
	
	private DateStyle(String style){
		this.style = style;
	}
	
	public String getStyle(){
		return this.style;
	}
}
