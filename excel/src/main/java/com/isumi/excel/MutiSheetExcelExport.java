package com.isumi.excel;

import com.google.common.collect.Maps;
import com.isumi.excel.annotations.ExcelEntity;
import com.isumi.excel.annotations.ExportField;
import com.isumi.excel.styles.Style;
import com.isumi.excel.utils.Reflections;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class MutiSheetExcelExport {



    /**
     * 导出excel
     * @param mutiLevelList 多级list，每个导出到一个sheet页
     * @param outStream 输出流
     */
    public void exportMutiSheetExcel(List<List> mutiLevelList, OutputStream outStream){
        try {
            HSSFWorkbook wb = new HSSFWorkbook();
            for(List dataList:mutiLevelList){

                if(!CollectionUtils.isEmpty(dataList)){
                    Class clazz = dataList.get(0).getClass();

                    Annotation annotation = clazz.getAnnotation(ExcelEntity.class);

                    ExcelEntity excelEntity = (ExcelEntity)annotation;
                    Class<?> styleClass = excelEntity.style();
                    Map<String, Map<String, Object>> fieldMap = Maps.newHashMap();
                    Map<Integer,String> fieldIndexMap = Maps.newHashMap();
                    parseExportField(clazz,excelEntity, fieldMap, fieldIndexMap);

                    Sheet sheet = wb.createSheet(excelEntity.sheetName());
                    Style style = (Style)(styleClass.newInstance());
                    CellStyle headStyle = style.getHeadStyle(wb);

                    Set<Map.Entry<Integer,String>> fieldIndexEs = fieldIndexMap.entrySet();
                    //创建标题栏
                    if(fieldIndexEs!=null&&!fieldIndexEs.isEmpty()){
                        Row row = sheet.createRow(0);// 创建第一行
                        row.setHeight(excelEntity.rowHeight());
//                        Cell cell = row.createCell(0);// 创建第一行的第一个单元格
//                        cell.setCellValue("序号");
//                        cell.setCellStyle(headStyle);
                        Cell cell;
                        Iterator<Map.Entry<Integer,String>> iterator = fieldIndexEs.iterator();
                        while(iterator.hasNext()){
                            Map.Entry<Integer,String> entry = iterator.next();
                            cell  = row.createCell(entry.getKey());
                            cell.setCellValue(entry.getValue());
                            cell.setCellStyle(headStyle);
                            Map<String,Object> ifm = fieldMap.get(entry.getValue());
                            short colWidth = (Short)(ifm.get("colWidth"));
                            sheet.setColumnWidth(entry.getKey(), colWidth);
                        }
                    }
                    //创建数据记录
                    if(dataList!=null&&!dataList.isEmpty()){
                        int rowIndex = 1;
                        CellStyle bodyStyle = style.getBodyStyle(wb);
                        if(excelEntity.wrapText()){
                            bodyStyle.setWrapText(excelEntity.wrapText());
                        }
                        for(Object data:dataList){
                            Row row = sheet.createRow(rowIndex);// 创建第一行
                            row.setHeight(excelEntity.rowHeight());
//                            Cell cell = row.createCell(0);//
//                            cell.setCellValue(rowIndex++);
//                            cell.setCellStyle(bodyStyle);
                            rowIndex++;
                            Cell cell;
                            Iterator<Map.Entry<Integer,String>> iterator = fieldIndexEs.iterator();
                            while(iterator.hasNext()){
                                Map.Entry<Integer,String> entry = iterator.next();
                                cell  = row.createCell(entry.getKey());
                                Map<String,Object> ifm = fieldMap.get(entry.getValue());
                                String fieldName = ObjectUtils.toString(ifm.get("fieldName"));
                                cell.setCellValue(ObjectUtils.toString(Reflections.getFieldValue(data, fieldName)));
                                cell.setCellStyle(bodyStyle);
                            }
                        }
                    }
                }
            }
            wb.write(outStream);
            outStream.flush();

        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            try {
                outStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    /**
     * @param excelEntity
     * @param fieldMap
     * @param fieldIndexMap
     */
    private void parseExportField(Class clazz,ExcelEntity excelEntity,
                                  Map<String, Map<String, Object>> fieldMap,
                                  Map<Integer, String> fieldIndexMap) {
        int index = 0;
        //获取需要导出的字段
        Field[] fields = clazz.getDeclaredFields();
        for (Field field:fields) {
            ExportField exportField = field.getAnnotation(ExportField.class);
            if(exportField!=null){
                Map<String,Object> exportFieldMap = Maps.newHashMap();
                exportFieldMap.put("colWidth", exportField.colWidth());
                exportFieldMap.put("fieldName", field.getName());
                fieldMap.put(exportField.colName(), exportFieldMap);
                if(excelEntity.sortHead()){
                    fieldIndexMap.put(exportField.index(), exportField.colName());
                }else{
                    fieldIndexMap.put(index, exportField.colName());
                    index++;
                }
            }
        }
    }


}
