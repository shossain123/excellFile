package excel.excel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {
		String file = "data/Sample.xls";
		FileOutputStream fos = new FileOutputStream(file);
		Workbook wb = new HSSFWorkbook();

		for (int i = 1; i <= 3; i = i + 1) {

			Sheet sh = wb.createSheet("Sheet "+i);
			Row r = sh.createRow(0);
			Row r1 = sh.createRow(1);
			Cell c = r.createCell(0);
			Cell c1 = r.createCell(2);
			Cell c2 = r1.createCell(2);

			c.setCellValue("java");
			c1.setCellValue("excell");
			c2.setCellValue("islam");
		}
		wb.write(fos);
		wb.close();
		fos.close();

	}

}
