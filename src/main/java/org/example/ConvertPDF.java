package org.example;

import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;

public class ConvertPDF {
    public static void generate(Workbook workbook, String pathFile) throws Exception {
        workbook.getConverterSetting().setSheetFitToPage(true);
        workbook.saveToFile(pathFile, FileFormat.PDF);
    }
}
