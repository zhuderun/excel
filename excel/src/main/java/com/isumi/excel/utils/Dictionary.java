package com.isumi.excel.utils;

import com.google.common.collect.Maps;
import org.apache.commons.lang3.StringUtils;

import java.util.Map;

/**
 * @author Administrator
 *
 */
public class Dictionary {
	private Map<String, Object> dictMap;

	public Dictionary(String dictJsonString) {
		dictMap = Maps.newHashMap();
		if (StringUtils.isNotEmpty(dictJsonString)) {
			dictMap = new JsonMapper().fromJson(dictJsonString, Map.class);
		}
	}

	public boolean containsKey(String key) {
		return dictMap.containsKey(key);
	}

	public Object getValue(String key) {
		return dictMap.get(key);
	}

}
