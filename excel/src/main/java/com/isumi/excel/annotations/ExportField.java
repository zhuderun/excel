package com.isumi.excel.annotations;


import org.apache.poi.ss.usermodel.Cell;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出字段属性配置
 * @author zdr
 *
 */
@Retention(RetentionPolicy.RUNTIME)     
@Target(ElementType.FIELD) 
public @interface ExportField {
	/**
	 * 导出的excel表格的列名称
	 * @return 列名
	 */
	String colName();
	
	/**
	 * excel表格字段值的类型，默认为String
	 * @return
	 */
	int cellType() default Cell.CELL_TYPE_STRING;
	
	/**
	 * excel列宽
	 * @return
	 */
	short colWidth() default 4570;
	
	boolean covert() default false;
	
	/**
	 * 导出的excel如果需要排序,则必须配置该属性，否则按照系统自动排序
	 * 导出的exel是否需要排序属性设置参照{@ExcelEntity}
	 * @return
	 */
	int index() default -1;
	

}
