package com.isumi.excel;


import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.isumi.excel.annotations.ImportField;
import com.isumi.excel.model.ImportModel;
import com.isumi.excel.utils.FieldValidator;
import com.isumi.excel.utils.Reflections;
import com.isumi.excel.utils.StringUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.isumi.excel.utils.Dictionary;

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.*;

/**
 * Excel导入
 * @author zdr
 * todo 校验列的唯一性
 */
public class ExcelImport<T>{
	Class<T> clazz;
	
	private List<String> errorMsg = Lists.newArrayList();
	
	/**
	 * 获取导入过程中产生的错误信息
	 * @return
	 */
	public List<String> getErrorMsg() {
		return errorMsg;
	}

	public ExcelImport(Class<T> clazz){
		this.clazz = clazz;
	}

	public Collection<T> importExcel(InputStream inputStream,String fileName, int headIndex, int bodyStart){
		Collection<T> dist = Lists.newArrayList();
		try {

			// 得到目标类的所有的字段列表
			Field field[] = clazz.getDeclaredFields();
			// 将所有标有Annotation的字段，也就是允许导入数据的字段,放入到一个map中
			Map<String,Map<String,Object>> fieldMap = Maps.newHashMap();
			// 循环读取所有字段
			parseImportField(field, fieldMap);
            Set<String> headMapValidat = new HashSet<>();
            headMapValidat.addAll(fieldMap.keySet());
			Sheet sheet = this.getSheet(inputStream,fileName, 0);
			if(sheet!=null){
				//获取列标题以及索引位置信息
				Row headRow = sheet.getRow(headIndex);
				Map<Integer,String> headMap = this.getHead(headRow);
				//遍历内容
				int maxColIndex = headRow.getLastCellNum();
				for(int index=bodyStart;index<=sheet.getLastRowNum();index++){
					Row row = sheet.getRow(index);
					if(row==null){
						errorMsg.add("第"+(index+1)+"行为空行");
						continue;
					}
					T object = clazz.newInstance();
					boolean rowValidateResult = true;
					FieldValidator fv = null;
					short minColIndex  = headRow.getFirstCellNum();
					//检查是否是全空白行 入过是全空白行，跳出此行
					boolean isAllEmptyRow = true;
					for(int colIndex=minColIndex; colIndex<maxColIndex; colIndex++) {
						Cell cell = row.getCell(colIndex);
						if(StringUtils.isNotEmpty(this.getStringCellValue(cell))){
							isAllEmptyRow = false;
							break;
						}
					}
					if(isAllEmptyRow){
						continue;
					}
					for(int colIndex=minColIndex; colIndex<maxColIndex; colIndex++) {
						Cell cell = row.getCell(colIndex);
						if(headMap.containsKey(colIndex)){
							if(fieldMap.containsKey(headMap.get(colIndex))){
								Map<String,Object> importFieldMap = fieldMap.get(headMap.get(colIndex));
                                headMapValidat.remove(headMap.get(colIndex));
								String fieldName  = ObjectUtils.toString(importFieldMap.get("fieldName"));
								fv = (FieldValidator)importFieldMap.get("fieldValidator");
								boolean cellValidateResult = fv.validate(index,colIndex,this.getStringCellValue(cell));
								if(!cellValidateResult){
									rowValidateResult = false;
									errorMsg.add(fv.getErrorMsg());
								}else{
									invokeSetter(object, cell, importFieldMap,fieldName);
								}
							}
						}
					}
					if(!headMapValidat.isEmpty()){
                        errorMsg.add("使用的模板与当前模板不一致");
                        return null;
                    }
					if(rowValidateResult){
						if(object instanceof ImportModel){
							ImportModel im =(ImportModel)object;
							im.setLineNumber(index+1);
						}
						dist.add(object);
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			errorMsg.add(e.getMessage());
			return null;
		}
		return dist;
	}
	
	/**
	 * 导入excel
	 * @param file 导入的文件名
	 * @param headIndex 表头索引行位置 从0开始
	 * @param bodyStart 内容索引行位置 从0开始
	 * @return
	 */
	public Collection<T> importExcel(File file,int headIndex,int bodyStart){
		Collection<T> dist = Lists.newArrayList();
        try {  

            // 得到目标类的所有的字段列表  
            Field field[] = clazz.getDeclaredFields();  
            // 将所有标有Annotation的字段，也就是允许导入数据的字段,放入到一个map中  
            Map<String,Map<String,Object>> fieldMap = Maps.newHashMap();  
            // 循环读取所有字段  
            parseImportField(field, fieldMap);             
            Sheet sheet = this.getSheet(file, 0);
            if(sheet!=null){
            	//获取列标题以及索引位置信息
            	Row headRow = sheet.getRow(headIndex);
            	Map<Integer,String> headMap = this.getHead(headRow);
            	//遍历内容
            	int maxColIndex = headRow.getLastCellNum();
            	for(int index=bodyStart;index<=sheet.getLastRowNum();index++){
            		Row row = sheet.getRow(index); 
            		T object = clazz.newInstance();  
            		boolean rowValidateResult = true;
            		FieldValidator fv = null;
            	    short minColIndex  = headRow.getFirstCellNum();
            	    for(int colIndex=minColIndex; colIndex<maxColIndex; colIndex++) {
            	    	Cell cell = row.getCell(colIndex);  
            			if(headMap.containsKey(colIndex)){
            				if(fieldMap.containsKey(headMap.get(colIndex))){
            					Map<String,Object> importFieldMap = fieldMap.get(headMap.get(colIndex));
            					String fieldName  = ObjectUtils.toString(importFieldMap.get("fieldName"));
            					fv = (FieldValidator)importFieldMap.get("fieldValidator");
            					boolean cellValidateResult = fv.validate(index,colIndex,this.getStringCellValue(cell));
            					if(!cellValidateResult){
            						rowValidateResult = false;
            						errorMsg.add(fv.getErrorMsg());
            					}else{
	            					invokeSetter(object, cell, importFieldMap,fieldName);
            					}
            				}
            			}
            		}
            		if(rowValidateResult){
            			if(object instanceof ImportModel){
            				ImportModel im =(ImportModel)object;
            				im.setLineNumber(index+1);
            			}
            			dist.add(object);
            		}
            	}
            }
        } catch (Exception e) {  
            e.printStackTrace();  
            errorMsg.add(e.getMessage());
            return null;  
        }  
		return dist;
	}

	/**
	 * @param field
	 * @param fieldMap
	 */
	private void parseImportField(Field[] field,
			Map<String, Map<String, Object>> fieldMap) {
		for (int i = 0; i < field.length; i++) {  
		    Field f = field[i];  
		    // 得到单个字段上的Annotation  
		    ImportField importField = f.getAnnotation(ImportField.class);
		    if (importField != null) { 
		    	Map<String,Object> importFieldMap = Maps.newHashMap();
		    	DataType[] dataTypes = importField.validate();
		    	if(dataTypes!=null){
		    		importFieldMap.put("fieldValidator", new FieldValidator(importField.colName(),dataTypes,importField.dateStyle(),importField.options(),importField.maxLength(),importField.mustLength()));
		    	}
		    	importFieldMap.put("fieldName", f.getName());
		    	importFieldMap.put("fieldType", f.getType());
		    	importFieldMap.put("dictionary", new Dictionary(importField.dictionary()));
		    	importFieldMap.put("dateStyle", importField.dateStyle().getStyle());
		    	fieldMap.put(importField.colName(), importFieldMap);
		    }  
		}
	}

	/**
	 * @param object
	 * @param cell
	 * @param importFieldMap
	 * @param fieldName
	 * @throws ParseException
	 */
	private void invokeSetter(T object, Cell cell,
			Map<String, Object> importFieldMap, String fieldName)
			throws ParseException {
		Class<?> classType =  (Class<?>) importFieldMap.get("fieldType");
		Dictionary dictionary = (Dictionary)importFieldMap.get("dictionary");
		String value =  getStringCellValue(cell);
		if(dictionary.containsKey(value)){
			value = ObjectUtils.toString(dictionary.getValue(value));
		}
		if(StringUtils.isNotEmpty(value)){
			if(classType==String.class)
			{
				Reflections.invokeSetter(object, fieldName,value);
			}
			else if(classType== Date.class)
			{
				String dateStyle = ObjectUtils.toString(importFieldMap.get("dateStyle"));
				Reflections.invokeSetter(object, fieldName, DateUtils.parseDate(value, dateStyle));
			}
			else if(classType==Boolean.class||classType==Boolean.TYPE)
			{
				Reflections.invokeSetter(object, fieldName, new Boolean(value));
			}
			else if(classType==Integer.class||classType==Integer.TYPE)
			{
				Reflections.invokeSetter(object, fieldName, new Integer(value));
			}
			else if(classType==Long.class||classType ==Long.TYPE)
			{
				Reflections.invokeSetter(object, fieldName, new Long(value));
			}

			else if(classType==Double.class||classType==Double.TYPE)
			{
				Reflections.invokeSetter(object, fieldName, new Double(value));
			}
			else if(classType==Float.class||classType==Float.TYPE)
			{
				Reflections.invokeSetter(object, fieldName, new Float(value));
			}
			else if(classType==BigDecimal.class)
			{
				Reflections.invokeSetter(object, fieldName, new BigDecimal(value));
			}
		}else{
			if(classType==String.class)
			{
				Reflections.invokeSetter(object, fieldName,value);
			}
		}
	}

	public Sheet getSheet(InputStream inputStream,String fileName,int index){
		// 将传入的File构造为FileInputStream;
		try {
			if(StringUtils.substringAfterLast(fileName, ".").indexOf("xlsx")!=-1) {
				XSSFWorkbook wb = new XSSFWorkbook(inputStream);
				XSSFSheet sheet = wb.getSheetAt(index);
				return sheet;
			}else {
				HSSFWorkbook book = new HSSFWorkbook(inputStream);
				HSSFSheet sheet = book.getSheetAt(index);
				return sheet;
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			errorMsg.add("Excel File is Not Found");
		} catch (IOException e) {
			e.printStackTrace();
			errorMsg.add("Excel File stream cannot be read");
		}finally{
			if(inputStream!=null){
				try {
					inputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return null;
	}
	
	private Sheet getSheet(File file,int index){
		// 将传入的File构造为FileInputStream;  
        FileInputStream in=null;
		try {
			in = new FileInputStream(file);
			if(StringUtils.substringAfterLast(file.getName(), ".").indexOf("xlsx")!=-1) {
				XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
				XSSFSheet sheet = wb.getSheetAt(index);
				return sheet;
			}else {
				HSSFWorkbook book = new HSSFWorkbook(in); 
		        HSSFSheet sheet = book.getSheetAt(index); 	       
		        return sheet;
			}	        
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			errorMsg.add("Excel File is Not Found");
		} catch (IOException e) {
			e.printStackTrace();
			errorMsg.add("Excel File stream cannot be read");
		}finally{
			if(in!=null){
				try {
					in.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return null;        
	}
	
	public List<String> getHead(File file,int index){
		Sheet sheet = this.getSheet(file, index);
		return getHead(sheet);		
	}
	
	public List<String> getHead(Sheet sheet){
		if(sheet!=null){
			Row row = sheet.getRow(0);
			Iterator<Cell> cellIterator =  row.cellIterator();
			List<String> heads = Lists.newArrayList();
			while(cellIterator.hasNext()){
				heads.add(this.getStringCellValue(cellIterator.next()));				
			}
			return heads;
		}
		return null;		
	}
	
	private Map<Integer,String> getHead(Row row){
		if(row!=null){
			Map<Integer,String> headMap = Maps.newHashMap();
			int minColIndex  = row.getFirstCellNum();
    	    int maxColIndex = row.getLastCellNum();
    	    for(int colIndex=minColIndex; colIndex<maxColIndex; colIndex++) {
    	    	Cell cell = row.getCell(colIndex);    
    	    	if(cell == null) {
    	    		continue;
    	    	}
    	    	headMap.put(cell.getColumnIndex(),StringUtils.replaceByRegex(getStringCellValue(cell), "\\s*|\t|\r|\n"));
    	    }
			return headMap;
		}
		return null;
	}
	
	/**
	 * 获取单元格数据内容为字符串类型的数据
	 * 
	 * @author
	 * @param cell
	 * @return
	 */
	public String getStringCellValue(Cell cell) {
		String strCell = "";
		if(cell==null)
			return strCell;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			strCell = cell.getRichStringCellValue().toString();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if(DateUtil.isCellDateFormatted(cell)){  
                Date date = cell.getDateCellValue() ;
                return DateFormatUtils.format(date,"yyyy-MM-dd") ;  
            }
			BigDecimal cellValue = new BigDecimal(String.valueOf(cell.getNumericCellValue()));
			strCell = cellValue.stripTrailingZeros().toPlainString();
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			strCell = ObjectUtils.toString(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_BLANK:
			strCell = "";
			break;
		default:
			strCell = "";
			break;
		}
		if (strCell.equals("") || strCell == null) {
			return "";
		}
		return strCell.trim();
	}
}
