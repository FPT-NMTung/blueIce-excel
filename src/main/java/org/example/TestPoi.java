package org.example;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;

public class TestPoi {
    public static void main(String[] args) throws Exception {
        File templateFile = new File("template-non-multi.xlsx");
        InputStream is = Files.newInputStream(templateFile.toPath());
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        XSSFWorkbook wb = new XSSFWorkbook(templateFile);
        Workbook workbook = StreamingReader.builder()
                .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
                .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                .open(is);            // InputStream or File for XLSX file (required)

        for (Sheet sheet : workbook){
            System.out.println(sheet.getSheetName());
            for (Row r : sheet) {
                for (Cell c : r) {
                    System.out.println(c.getStringCellValue());
                }
            }
        }

//        SXSSFWorkbook resultWb = new SXSSFWorkbook(wb);
//        SXSSFSheet sheet = resultWb.getSheetAt(0);
//        SXSSFRow row = sheet.getRow(1);
//        SXSSFCell cell = row.getCell(0);
//
//        cell.setCellValue("123123123123");



    }
}
