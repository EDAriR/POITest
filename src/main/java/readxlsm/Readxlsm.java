package readxlsm;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Readxlsm {


    public static void main(String[] args) {

        String fileName = "new_file.xlsm";


        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(
                    OPCPackage.open("report1.xlsm")
            );
        } catch (IOException e1) {
            e1.printStackTrace();
        } catch (InvalidFormatException e1) {
            e1.printStackTrace();
        }

        Sheet sheet = workbook.getSheetAt(11);

        Row row = sheet.getRow(0);

        for (int i = 0; i < row.getLastCellNum(); i++) {

            System.out.println(row.getCell(i));
        }

        for (int i = 0; i < sheet.getLastRowNum(); i++) {

            System.out.println(sheet.getRow(i).getCell(0));
            Cell cell = sheet.getRow(i).getCell(2);

            cell.setCellValue(55);
        }


        //DO STUF WITH WORKBOOK

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(new File(fileName));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        try {
            workbook.write(out);
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        try {
            out.close();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        System.out.println("xlsm created successfully..");

    }
}
