package com.isumi.excel.annotations;




import com.isumi.excel.DataType;
import com.isumi.excel.DateStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;



/**
 * 导入字段属性配置
 * @author zdr
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ImportField {
	/**
	 * 导入的excel表格的列名称
	 * @return 列名
	 */
	String colName();

	/**
	 * excel列对应的需要校验的数据类型,默认为{}
	 * @return
	 */
	DataType[] validate() default {};

	/**
	 * excel列值允许出现的可选值(类似数据有效性中的下拉选项)
	 * @return
	 */
	String[] options() default {};

	/**
	 * 当指定数据类型为Date时，需指定日期的格式, 默认为YYYYMMDD
	 * @return
	 */
	DateStyle dateStyle() default DateStyle.YYYYMMDD;

	/**
	 * excel列对应的数据字典值,导入的时候将值转换为其对应的字典值
	 * 例如：excel性别列的值为男，对应的字典值为01，那么导入后返回的对象的字段值为01
	 * @return
	 */
	String dictionary() default "";

	int maxLength() default -1;

	int mustLength() default  -1;


}
