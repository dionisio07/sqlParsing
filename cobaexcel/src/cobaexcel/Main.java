import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author dio
 */
public class Main {

    /**
     * @param args the command line arguments
     */
    public static void main(String[]args) {
        try {

/*Nama file excell*/
            String filename = "D:/FileExcell.xls" ;
            HSSFWorkbook workbook = new HSSFWorkbook();

/*menentukan sheet*/
            HSSFSheet sheet = workbook.createSheet("FirstSheet"); 

            HSSFRow rowhead = sheet.createRow((short)0);  /*index row 0*/
                    rowhead.createCell(0).setCellValue("No");  /*column 0*/
                    rowhead.createCell(1).setCellValue("Nama");  /*column 1*/
                    rowhead.createCell(2).setCellValue("Alamat");  /*column 2*/

            HSSFRow row = sheet.createRow((short)1);   /*index row 1*/
                    row.createCell(0).setCellValue("1");  /*column 0*/
                    row.createCell(1).setCellValue("Okin Luberto");  /* column  1*/
                    row.createCell(2).setCellValue("Indonesia");  /* column 2*/
                   
            HSSFRow row2 = sheet.createRow((short)2);   /*index row 2*/
                    row2.createCell(0).setCellValue("2"); /*column 0*/
                    row2.createCell(1).setCellValue("Ayrini"); /* column  1*/
                    row2.createCell(2).setCellValue("Indonesia"); /* column 2*/

            FileOutputStream fileOut = new FileOutputStream(filename);

/*menulis file*/
            workbook.write(fileOut);

/*menutup koneksi*/
            //System.out.println(row.getCell(1));
            fileOut.close();
            System.out.println(row.getCell(1));
            System.out.println("Excel berhasil di buat !");
        } catch ( Exception ex ) {
            System.out.println(ex);
        }
    }
}
