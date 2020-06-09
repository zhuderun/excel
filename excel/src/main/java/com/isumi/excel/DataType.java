package com.isumi.excel;


/**
 * 字段数据类型
 * @author zdr
 *
 */
public enum DataType {
	/**
	 * 日期
	 */
	DATE("date"),
	/**
	 * EMAIL
	 */
	EMAIL("email"),
	/**
	 * 身份证号
	 */
	ID_CARD("idCard"),
	/**
	 * 手机号
	 */
	MOBILE_PHONE("mobilePhone"),
	/**
	 * 数字(只包含数字，或者是空，有小数点、正负号都不行)
	 */
	NUMERIC("numeric"),
	/**
	 * 必填
	 */
	REQURIED("required"),
	/**
	 * 下拉选择
	 */
	OPTIONS("options"),

	/**
	 * 数值型，可以是正负数、小数，当为空串时不验证
	 */
	NUMBER("number"),
	/**
	 * 最大长度，配合maxlength使用
	 */
	MAXLENGTH("maxLength"),

	WEEKIDCARD("weekIdCard"),

	MUSTLENGTH("mustLength");
	
	private String name;
	
	private DataType(String name){
		this.name = name;
	}
	
	public String getName(){
		return this.name;
	}
	
}
