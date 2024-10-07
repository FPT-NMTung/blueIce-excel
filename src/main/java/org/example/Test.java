package org.example;

import com.spire.xls.InsertMoveOption;
import com.spire.xls.InsertOptionsType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

public class Test {
    public static void main(String[] args) {
        long totalTime;
        long startTime;
        System.out.print("Read file ... ");
        totalTime = System.currentTimeMillis();
        startTime = System.currentTimeMillis();
        Workbook workbook = new Workbook();
        workbook.loadFromFile("template-non-multi.xlsx");
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        Worksheet worksheet = workbook.getWorksheets().get(0);
//        worksheet.();

        startTime = System.currentTimeMillis();
        System.out.print("Export 1 ... ");
        workbook.saveToFile("result.xlsx");
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        System.out.println((System.currentTimeMillis() - totalTime) + "ms");
    }
}
