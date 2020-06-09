package com.isumi.excel.compress;



import com.isumi.excel.compress.zip.ZIPReader;
import com.isumi.excel.compress.zip.ZIPWriter;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;

public class Compress {
    public static ZIPReader unzip(File file) {
        return new ZIPReader(file);
    }

    public static ZIPWriter zip() {
        return new ZIPWriter();
    }

    public static void main(String[] args) {
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        wb.write(out);
//        ByteArrayInputStream in = new ByteArrayInputStream(out.toByteArray());
//        zw.addFile(fileName,in);
//        response.setCharacterEncoding("GBK");
//        response.setContentType("application/xls;charset=GBK");
//        String outName = new String("压缩.zip".getBytes("GBK"),"ISO-8859-1");
//        response.setHeader("Content-disposition", "attachment; filename="+ outName);
//        zw.level(9).write(response.getOutputStream());
    }

}
