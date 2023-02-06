package excel.excel;


import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelwrite2 {

	public static void main(String[] args) throws IOException {
		String file = "data/sample2.xlsx";
		FileOutputStream fos = new FileOutputStream (file);
		Workbook wb = new XSSFWorkbook ();
		
		for(int i=1; i<3; i++) {
		Sheet sh = wb.createSheet("Sheet"+i);
		 Row     r   =  sh.createRow(0);
		 Row r1 = sh.createRow(5);
		 Cell       c    =    r.createCell(0);
		  Cell c1 = r.createCell(2) ;
		  Cell c2 = r1.createCell(5);
		  
		 c.setCellValue("rahim");
		 c1.setCellValue("kuddus");
		 c2.setCellValue("saddam");
		}
		 wb.write(fos);
		 wb.close();
		 fos.close();
		

	}

}
