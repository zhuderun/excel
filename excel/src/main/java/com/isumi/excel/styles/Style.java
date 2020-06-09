package com.isumi.excel.styles;


import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author zdr
 *
 */
public abstract class Style {
	

	/**
	 * 获取表头样式
	 * @param workbook
	 * @return
	 */
	public  abstract  CellStyle getHeadStyle(Workbook workbook);
	
	/**
	 * 获取表格内容样式
	 * @param workbook
	 * @return
	 */
	public  abstract  CellStyle getBodyStyle(Workbook workbook);


}
