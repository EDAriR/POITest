package chart;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class OOXMLChart {

    /**
     * 修改模板数据并保存
     *
     * @author cweili
     *
     * @param titles
     *            标题行
     * @param values
     *            数值
     * @param inFile
     *            模板文件
     * @param outFile
     *            输出文件
     */
    public static void createChart(String[] titles, double[] values, String inFile,
                                   String outFile) {
        try {
            // 读取模板
            FileInputStream is = new FileInputStream(inFile);
            Workbook wb = WorkbookFactory.create(is);
            // 读取工作表0
            Sheet sheet0 = wb.getSheetAt(0);
            System.out.println("行数: " + (sheet0.getLastRowNum() + 1));
            Row titleRow = sheet0.getRow(0);
            for (int i = 0; i < titleRow.getLastCellNum(); ++i) {
                Cell cell = titleRow.getCell(i);
                cell.setCellValue(titles[i]);
            }
            // 数据项目
            Row row = sheet0.getRow(1);
            System.out.println("列数: " + row.getLastCellNum());
            for (int i = 0; i < row.getLastCellNum(); ++i) {
                Cell cell = row.getCell(i);
                cell.setCellValue(values[i]);
            }
            // 输出文件
            FileOutputStream os = new FileOutputStream(outFile);
            wb.write(os);
            is.close();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
