package readxlsx;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;

public class ReadExcel {
    /**
     * 读取Excel测试，兼容 Excel 2003/2007/2010
     */
    public static void main(String[] args) {

//    }
//    public String readExcel()
        {
            SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
            try {
                //同时支持Excel 2003、2007
                File excelFile = new File("report1.xlsm"); //创建文件对象
                FileInputStream is = new FileInputStream(excelFile); //文件流
                Workbook workbook = WorkbookFactory.create(is); //这种方式 Excel 2003/2007/2010 都是可以处理的
                int sheetCount = workbook.getNumberOfSheets();  //Sheet的数量
                //遍历每个Sheet
                for (int s = 0; s < sheetCount; s++) {

                    System.out.println("====================");
                    System.out.println("read sheet : " + s);
                    System.out.println("====================");

                    Sheet sheet = workbook.getSheetAt(s);
                    int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数
                    //遍历每一行
                    for (int r = 0; r < rowCount; r++) {
                        Row row = sheet.getRow(r);
                        int cellCount = row.getPhysicalNumberOfCells(); //获取总列数
                        //遍历每一个单元格
                        for (int c = 0; c < cellCount; c++) {
                            Cell cell = row.getCell(c);
                            int cellType = cell.getCellType();
                            String cellValue = null;

                            //在读取单元格内容前,设置所有单元格中内容都是字符串类型
                            cell.setCellType(Cell.CELL_TYPE_STRING);

                            //按照字符串类型读取单元格内数据
                            cellValue = cell.getStringCellValue();

                    /*在这里可以对每个单元格中的值进行二次操作转化*/

                            System.out.print(cellValue + "    ");
                        }
                        System.out.println();
                    }
                }

            } catch (Exception e) {
                e.printStackTrace();
            }

//        return Action.SUCCESS;
//    }
        }
    }
}
