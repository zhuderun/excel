package com.isumi.excel.utils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Validator {
	
	public static boolean checkEmail(String mail) {
		String regex = "\\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(mail);
		return m.find();
	}

	public static boolean validate11Mobile(String phone) {
		int l = phone.length();
		boolean rs = false;
		switch (l) {
			case 11:
				if (matchingText("^(1[3-9]\\d{9}$)", phone)) {
					rs = true;
				}
				break;
			default:
				rs = false;
				break;
		}
		return rs;
	}


	public static boolean validateMobile(String phone) {
		int l = phone.length();
		boolean rs = false;
		switch (l) {
		case 7:
			if (matchingText("^(13[0-9]|15[0-9]|18[7|8|9|6|5])\\d{4}$", phone)) {
				rs = true;
			}
			break;
		case 11:
			if (matchingText("^(1[3-9]\\d{9}$)", phone)) {
				rs = true;
			}
			break;
		default:
			rs = false;
			break;
		}
		return rs;
	}

	public static boolean validateWeekCardNo(String weekCardNo){
		boolean rs = false;
		if (matchingText("^([A-Za-z0-9]+$)", weekCardNo)) {
			rs = true;
		}
		return rs;
	}

	public static boolean matchingText(String expression, String text) {
		Pattern p = Pattern.compile(expression); // 姝ｅ垯琛ㄨ揪寮?
		Matcher m = p.matcher(text); // 鎿嶄綔鐨勫瓧绗︿覆
		boolean b = m.matches();
		return b;
	}
	
	public static boolean validateIdCards(String idCards){
		return IdCardUtils.validateCard(idCards);
	}

}
