package com.isumi.excel.utils;


import com.isumi.excel.DataType;
import com.isumi.excel.DateStyle;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateUtils;

import java.text.ParseException;

/**
 * 字段类型校验工具类
 * @author Administrator
 *
 */
public class FieldValidator {
	
	private DataType[] dataTypes;
	
	private DateStyle dateStyle;
	
	private String errorMsg;
	
	private String colName;
	
	private String[] options;

	private int maxLength;

	private int mustLength;
	
	public String getErrorMsg() {
		return errorMsg;
	}

	public FieldValidator(String colName,DataType[] dataTypes,DateStyle dateStyle,String[] options,int maxLength,int mustLenfth){
		this.colName = colName;
		this.dataTypes = dataTypes;
		this.dateStyle = dateStyle;
		this.options = options;
		this.maxLength = maxLength;
		this.mustLength = mustLenfth;
	}
	
	public boolean validate(int row,int col,String fieldValue){
		boolean flag = true;
		errorMsg = null;
		if(dataTypes!=null){
			StringBuffer sb = new StringBuffer((row+1)+"行"+(col+1)+"列:");
			for(DataType dataType:dataTypes){
				if(dataType==DataType.DATE){
					String[] args = {fieldValue,dateStyle.getStyle()};
					boolean result = (Boolean)(Reflections.invokeMethodByName(this,"validate"+StringUtils.capitalize(dataType.getName()), args));
					if(!result){
						flag = false;
						sb.append(colName+" 不合法  !");
					}
				}else{
					String[] args = {fieldValue};
					boolean result = (Boolean)Reflections.invokeMethodByName(this,"validate"+StringUtils.capitalize(dataType.getName()), args);
					if(!result){
						flag = false;
						if(dataType==DataType.REQURIED){
							sb.append(colName+" 必填！");
							break;
						}else if(dataType==DataType.OPTIONS){
							sb.append(colName+" 必须在选项范围内：").append(ArrayUtils.toString(options));
						}else if (dataType==DataType.MAXLENGTH){
							sb.append(colName+"字符长度必须在").append(maxLength).append("个以内");
						}else if(dataType==DataType.MUSTLENGTH){
							sb.append(colName+"字符长度必须等于").append(mustLength);
						}
						else{
							sb.append(colName+" 不合法  !");
						}
					}
				}
				
			}
			if(!flag){
				errorMsg = sb.toString();
			}
		}
		return flag;
	}
	
	/**
	 * 验证日期
	 * @param date
	 * @param dateStyle
	 * @return
	 */
	public  boolean validateDate(String date,DateStyle dateStyle){
		if(date==null||dateStyle==null){
			return false;
		}else if(date.length()!=dateStyle.getStyle().length()){
			return false;
		}else{
			try {
				DateUtils.parseDate(date, dateStyle.getStyle());
			} catch (ParseException e) {
				e.printStackTrace();
				return false;
			}
		}
		return true;
	}
	
	/**
	 * 验证email地址
	 * @param email
	 * @return
	 */
	public  boolean validateEmail(String email){
		return Validator.checkEmail(email);
	}
	
	/**
	 * 验证身份证号 
	 * @param idCard
	 * @return
	 */
	public boolean validateIdCard(String idCard){
		return Validator.validateIdCards(idCard);
	}
	
	/**
	 * 验证手机号
	 * @param mobilePhone
	 * @return
	 */
	public  boolean validateMobilePhone(String mobilePhone){
		return Validator.validateMobile(mobilePhone);
	}
	
	/**
	 * 验证是否数字
	 * @param numeric
	 * @return
	 */
	public  boolean validateNumeric(String numeric){
		if(StringUtils.isEmpty(numeric)){
			return true;
		}
		return StringUtils.isNumeric(numeric);
	}
	
	/**
	 * 验证是否必填
	 * @param required
	 * @return
	 */
	public  boolean validateRequired(String required){
		return StringUtils.isNotEmpty(required);
	}
	
	/**
	 * 验证Options
	 * @param option
	 * @return
	 */
	public boolean validateOptions(String option){
		if(StringUtils.isEmpty(option)){
			return true;//为空不验证
		}
		if(option!=null){
			return ArrayUtils.contains(options, option);
		}
		return false;
	}

	/**
	 * 验证数值
	 * @return
	 */
	public boolean validateNumber(String number){
		boolean result = true;
		if(StringUtils.isNotEmpty(number)){
			result =  NumberUtils.isNumber(number);
		}
		return result;
	}

	public boolean validateMaxLength(String value){
		boolean result = true;
		if(StringUtils.isNotEmpty(value)&&maxLength>0){
			return value.length()<=maxLength;
		}
		return result;
	}

	public boolean validateWeekIdCard(String value){
		boolean result = true;
		if(StringUtils.isNotEmpty(value)){
			result = Validator.validateWeekCardNo(value);
		}
		return result;
	}

	public boolean validateMustLength(String value){
		boolean result = true;
		if(StringUtils.isNotEmpty(value)&&mustLength>0){
			result = value.length() == mustLength;
		}
		return result;
	}
}
