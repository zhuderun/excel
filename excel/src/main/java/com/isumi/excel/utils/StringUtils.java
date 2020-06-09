package com.isumi.excel.utils;

import org.apache.commons.lang3.time.DateFormatUtils;

import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Description: 字符串操作工具类 继承了apache-commons-lang3包中的StringUtils类 增加了时间转换为mills字符串
 * mills字符串转换为时间等相关方法 <br>
 * 
 * @author 梁华璜
 */
public class StringUtils extends org.apache.commons.lang3.StringUtils {
	// Constants used by escapeHTMLTags
	private static final char[] QUOTE_ENCODE = "&quot;".toCharArray();
	private static final char[] AMP_ENCODE = "&amp;".toCharArray();
	private static final char[] LT_ENCODE = "&lt;".toCharArray();
	private static final char[] GT_ENCODE = "&gt;".toCharArray();
	private static final String regEx_html = "<[^>]+>";

	/**
	 * This method takes a string which may contain HTML tags (ie, &lt;b&gt;,
	 * &lt;table&gt;, etc) and converts the '&lt'' and '&gt;' characters to
	 * their HTML escape sequences. It will also replace LF with &lt;br&gt;.
	 * 
	 * @param in
	 *            the text to be converted.
	 * @return the input string with the characters '&lt;' and '&gt;' replaced
	 *         with their HTML escape sequences.
	 */
	public static String escapeHTMLTags(String in) {
		if (in == null) {
			return null;
		}
		char ch;
		int i = 0;
		int last = 0;
		char[] input = in.toCharArray();
		int len = input.length;
		StringBuilder out = new StringBuilder((int) (len * 1.3));
		for (; i < len; i++) {
			ch = input[i];
			if (ch > '>') {
			} else if (ch == '<') {
				if (i > last) {
					out.append(input, last, i - last);
				}
				last = i + 1;
				out.append(LT_ENCODE);
			} else if (ch == '>') {
				if (i > last) {
					out.append(input, last, i - last);
				}
				last = i + 1;
				out.append(GT_ENCODE);
			} else if (ch == '\n') {
				if (i > last) {
					out.append(input, last, i - last);
				}
				last = i + 1;
				out.append("<br>");
			}
		}
		if (last == 0) {
			return in;
		}
		if (i > last) {
			out.append(input, last, i - last);
		}
		return out.toString();
	}

	/**
	 * 除去字符串中的空格、回车、换行符、制表符
	 * 
	 * @param str
	 * @return
	 */
	public static String replaceBlank(String str) {
		Pattern p = Pattern.compile("\\s*|\t|\r|\n");
		Matcher m = p.matcher(str);
		String after = m.replaceAll("");
		return after;
	}
	
	public static String replaceByRegex(String str,String regex){
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(str);
		String after = m.replaceAll("");
		return after;
	}
	
	/**
	 * 过滤html标签
	 * @param str
	 * @return
	 */
	public static String replaceHtml(String str){
		Pattern p = Pattern.compile(regEx_html, Pattern.CASE_INSENSITIVE);
        Matcher m = p.matcher(str);
        String after = m.replaceAll("");
		return after;
	}

	/**
	 * js日历控件日期转换为字符串类型的mills
	 * 
	 * @param jsCalendar
	 *            js日期 支持格式为20100520,2010-05-20
	 * @param separator
	 *            分隔符 如格式为20100520 则为空 格式为2010-05-20则为"-"
	 * @return 如果jsCalendar为空则返回空字符串
	 */
	public static Date jsCalendarToDate(String jsCalendar, String separator) {
		Calendar cal = Calendar.getInstance();
		if (StringUtils.isBlank(jsCalendar)) {
			return null;
		}
		if (separator != null && !separator.trim().equals("")) {
			String[] jsCalendarArr = StringUtils.split(jsCalendar, separator);
			cal.set(Integer.parseInt(jsCalendarArr[0]),
					Integer.parseInt(jsCalendarArr[1]) - 1,
					Integer.parseInt(jsCalendarArr[2]));
		} else {
			cal.set(Integer.parseInt(StringUtils.substring(jsCalendar, 0, 4)),
					Integer.parseInt(StringUtils.substring(jsCalendar, 4, 6)) - 1,
					Integer.parseInt(StringUtils.substring(jsCalendar, 6, 8)));
		}
		return cal.getTime();
	}

	/**
	 * 字符串转换成日历控件日期
	 * 
	 * @param pattern
	 * @return
	 */
	public static String dateToJsCalendar(Date date, String pattern) {
		return DateFormatUtils.format(date, pattern);
	}

	public static boolean isNumeric(String num) {
		try {
			Double.parseDouble(num);
			return true;
		} catch (NumberFormatException e) {
			return false;
		}
	}

	public static String intToChnNumConverter(int num) {
		String resultNumber = null;
		if (num > 10000 || num < 0) {
			return "";
		}
		HashMap<Integer, String> chnNumbers = new HashMap<Integer, String>();
		chnNumbers.put(0, "零");
		chnNumbers.put(1, "一");
		chnNumbers.put(2, "二");
		chnNumbers.put(3, "三");
		chnNumbers.put(4, "四");
		chnNumbers.put(5, "五");
		chnNumbers.put(6, "六");
		chnNumbers.put(7, "七");
		chnNumbers.put(8, "八");
		chnNumbers.put(9, "九");

		HashMap<Integer, String> unitMap = new HashMap<Integer, String>();
		unitMap.put(1, "");
		unitMap.put(10, "十");
		unitMap.put(100, "百");
		unitMap.put(1000, "千");
		int[] unitArray = { 1000, 100, 10, 1 };

		StringBuilder result = new StringBuilder();
		int i = 0;
		while (num > 0) {
			int n1 = num / unitArray[i];
			if (n1 > 0) {
				result.append(chnNumbers.get(n1)).append(
						unitMap.get(unitArray[i]));
			}
			if (n1 == 0) {
				if (result.lastIndexOf("零") != result.length() - 1) {
					result.append("零");
				}
			}
			num = num % unitArray[i++];
			if (num == 0) {
				break;
			}
		}
		resultNumber = result.toString();
		if (resultNumber.startsWith("零")) {
			resultNumber = resultNumber.substring(1);
		}
		if (resultNumber.startsWith("一十")) {
			resultNumber = resultNumber.substring(1);
		}
		return resultNumber;
	}

}
