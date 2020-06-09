package com.isumi.excel.annotations;




import com.isumi.excel.styles.StandardStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


/**
 * 凡需要进行excel导入导出的实体对象，都需要配置注解信息
 * @author zdr
 *
 */
@Retention(RetentionPolicy.RUNTIME)     
@Target(ElementType.TYPE) 
public @interface ExcelEntity {
	/**
	 * 导出excel表格时候配置的表格样式，默认为StandardStyle，如需要更改需自定义实现了Style接口的样式
	 * @return
	 */
	Class<?> style() default StandardStyle.class;
	
	/**
	 * 默认导出的excel表格工作簿的名字，默认值为sheet0
	 * @return
	 */
	String sheetName() default "sheet0";
	
	/**
	 * 导出excel表格时的保留字段属性，为以后在表格顶部增加统一标题预留
	 * @return
	 */
	String headTitle() default "";
	
	/**
	 * 导出excel表格时，是否需要排序表头字段,默认为false不排序
	 * @return
	 */
	boolean sortHead() default false;
	
	/**
	 * 导出excel表格时的行高，默认为470
	 * @return
	 */
	short rowHeight() default 470;
	
	
	/**
	 * 字段值是否需要自动换行，默认为false不自动换行
	 * @return
	 */
	boolean wrapText() default false;

	short headRowHeight() default 470;

	short headTitleRowHeight() default 470;
}
