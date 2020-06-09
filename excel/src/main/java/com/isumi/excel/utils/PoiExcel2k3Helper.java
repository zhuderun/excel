package com.isumi.excel.utils;

import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel 读取（97-2003格式）
 * @author	chengesheng
 * @date	2012-4-27 下午03:39:01
 * @note	PoiExcel2k3Helper
 */
public class PoiExcel2k3Helper extends PoiExcelHelper {
	/** 获取sheet列表 */
	public ArrayList<String> getSheetList(File file) {
		ArrayList<String> sheetList = Lists.newArrayListWithCapacity(0);
		try {
			HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
			int i = 0;
			while (true) {
				try {
					String name = wb.getSheetName(i);
					sheetList.add(name);
					i++;
				} catch (Exception e) {
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return sheetList;
	}

	/** 读取Excel文件内容 */
	public List<List<String>> readExcel(File file, int sheetIndex, String rows, String columns) {
		List<List<String>> dataList = Lists.newArrayList();
		try {
			HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
			HSSFSheet sheet = wb.getSheetAt(sheetIndex);			
			dataList = readExcel(sheet, rows, getColumnNumber(sheet, columns));
		} catch (Exception e) {
			e.printStackTrace();
		}
		return dataList;
	}
	
	/** 读取Excel文件内容 */
	public List<List<String>> readExcel(File file, int sheetIndex, String rows, int[] cols) {
		List<List<String>> dataList = Lists.newArrayList();
		try {
			HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
			HSSFSheet sheet = wb.getSheetAt(sheetIndex);
			dataList = readExcel(sheet, rows, cols);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return dataList;
	}
}