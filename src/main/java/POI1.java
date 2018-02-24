import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class POI1 {

//    New Workbook
//HSSF is the POI Project's pure Java implementation of the Excel '97(-2007) file format.
//
//        Workbook wb = new HSSFWorkbook();
//
//        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
//
//        wb.write(fileOut);
//        fileOut.close();



// XSSF is the POI Project's pure Java implementation of the Excel 2007 OOXML (.xlsx) file format.
    Workbook wb = new XSSFWorkbook();
//    FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
//
//    wb.write(fileOut);
//    fileOut.close();
}
