package com.isumi.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.math.BigDecimal;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * excel导出工具类
 * 
 * @author guojie
 * 
 */
public class ExcelUtils {

    /**
     * 导出excel
     * 
     * @param headers
     *            从外到里，第一层list表示几个头标题，第二层list表示每一行的标题的名称和占用几个单元格，如果需要设置，则用加上count元素，不设置默认占用一个单元格
     * @param bodyContent
     */
    public static HSSFWorkbook exportExcel(List<List<Map<String, Object>>> headers, List<Object[]> bodyContent) {
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 产生工作表对象
            HSSFSheet sheet = workbook.createSheet();

            // sheet表、单元格样式、数据集合
            createExcelHeader(workbook, sheet, headers);

            // sheet表、body的数据集合
            createExcelBody(workbook, sheet, bodyContent, headers.size());

            return workbook;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 创建excel表头
     * 
     * @param workbook
     * @param sheet
     * @param lists
     */
    private static void createExcelHeader(HSSFWorkbook workbook, HSSFSheet sheet,
        List<List<Map<String, Object>>> lists) {
        // 设置第一个工作表的名称为firstSheet
        workbook.setSheetName(0, "sheet");
        sheet.setDefaultColumnWidth(25);// 设置默认每一列的宽度
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 设置粗体
        cellStyle.setFont(font);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        for (int i = 0; i < lists.size(); i++) {// 行
            HSSFRow row = sheet.createRow(i);
            HSSFCell cell = null;
            int fromIndex = 0; // 合并单元格的起始列
            List<Map<String, Object>> innerList = lists.get(i);
            for (int j = 0; j < innerList.size(); j++) {// 列
                cell = row.createCell(fromIndex);
                Map<String, Object> map = innerList.get(j);
                if (map.get("count") == null) {// 合并单元格数量
                    cell.setCellValue((String)map.get("headName"));
                    cell.setCellStyle(cellStyle);
                    fromIndex++;
                } else {
                    cell.setCellValue((String)map.get("headName"));
                    cell.setCellStyle(cellStyle);
                    int count = fromIndex + (Integer)map.get("count") - 1;
                    // 合并单元格
                    CellRangeAddress region1 = new CellRangeAddress(i, i, (short)fromIndex, (short)count); // 参数1：起始行
                                                                                                           // 参数2：终止行
                                                                                                           // 参数3：起始列
                                                                                                           // 参数4：终止列
                    sheet.addMergedRegion(region1);
                    fromIndex = count + 1;
                }
            }
        }
    }

    /**
     * 创建excel表格数据
     * 
     * @param workbook
     * @param sheet
     * @param list
     * @param size
     */
    private static void createExcelBody(HSSFWorkbook workbook, HSSFSheet sheet, List<Object[]> list, Integer size) {
        // 定义单元格样式
        HSSFCellStyle cs = workbook.createCellStyle();
        cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 对齐方式
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.BLUE.index);
        cs.setFont(font);
        HSSFRow row;
        HSSFCell cell;
        if (list.size() > 0) {
            for (int i = 0; i < list.size(); i++) {
                Object[] arr = list.get(i);

                row = sheet.createRow(i + size);

                int index = 0;
                for (Object o : arr) {
                    cell = row.createCell(index++);
                    if (o instanceof String) {
                        HSSFDataFormat format = workbook.createDataFormat();
                        cs.setDataFormat(format.getFormat("@"));
                        cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                        cell.setCellStyle(cs);
                        cell.setCellValue(isNullOrEmpty(o) ? "" : o.toString());
                    } else if (o instanceof Integer) {
                        cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                        cell.setCellStyle(cs);
                        Integer value = (Integer)o;
                        cell.setCellValue(String.valueOf(value));
                    } else if (o instanceof Double) {
                        cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                        cell.setCellStyle(cs);
                        Double number = (double)o;
                        if(number%1==0){
                            BigDecimal b = new BigDecimal(number);
                            String value = String.valueOf(b.intValue());
                            cell.setCellValue(value);
                        }else{
                            BigDecimal b = new BigDecimal(number);
                            String value = String.valueOf(b.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue());
                            cell.setCellValue(value);
                        }
                    } else if (o instanceof Long) {
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        cell.setCellStyle(cs);
                        cell.setCellValue(String.valueOf((Long)o));
                    } else if (o instanceof BigDecimal) {
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        cell.setCellStyle(cs);
                        BigDecimal b = (BigDecimal)o;
                        if(new BigDecimal(b.intValue()).compareTo(b)==0){//说明是整数
                            String value = String.valueOf(b.intValue());
                            cell.setCellValue(value);
                        } else {
                            String value = String.valueOf(b.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue());
                            cell.setCellValue(value);
                        }
                    }
                }
            }
        }
    }

    public static boolean isNullOrEmpty(Object obj) {
        boolean isEmpty = false;
        if (obj == null) {
            isEmpty = true;
        } else if (obj instanceof String) {
            isEmpty = ((String)obj).trim().isEmpty();
        } else if (obj instanceof Collection) {
            isEmpty = (((Collection)obj).size() == 0);
        } else if (obj instanceof Map) {
            isEmpty = ((Map)obj).size() == 0;
        } else if (obj.getClass().isArray()) {
            isEmpty = Array.getLength(obj) == 0;
        }
        return isEmpty;
    }
}
