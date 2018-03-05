import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Images {

    public static void main(String[] args) {
        //create a new workbook
        Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();

        //add picture data to this workbook.
        InputStream is = null;
        try {
            is = new FileInputStream("E8zwVe1PWj.jpeg");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        byte[] bytes = new byte[0];
        try {
            bytes = IOUtils.toByteArray(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        try {
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        CreationHelper helper = wb.getCreationHelper();

        //create sheet
        Sheet sheet = wb.createSheet();

        // Create the drawing patriarch.  This is the top level container for all shapes.
        Drawing drawing = sheet.createDrawingPatriarch();

        //add a picture shape
        ClientAnchor anchor = helper.createClientAnchor();
        //set top-left corner of the picture,
        //subsequent call of Picture#resize() will operate relative to it
        anchor.setCol1(10);
        anchor.setRow1(10);
        Picture pict = drawing.createPicture(anchor, pictureIdx);

        //auto-size picture relative to its top-left corner
        pict.resize();

        //save workbook
        String file = "picture.xls";
        if(wb instanceof XSSFWorkbook) file += "x";
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(file);
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
