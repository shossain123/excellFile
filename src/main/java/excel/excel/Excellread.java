package excel.excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class Excellread {

	public static void main(String[] args) throws IOException {
		String file = "data/sample.xls";
		FileInputStream fis = new FileInputStream (file);
		Workbook wb = new HSSFWorkbook(fis);
		Sheet sh = wb.getSheet("Sheet1");
//		int totalRow = sh.getPhysicalNumberOfRows();
//		int cell = sh.getRow(0).getPhysicalNumberOfCells();
//		System.out.println(totalRow);
//		System.out.println(cell);
		for (Row r: sh) {
			
			for(Cell c: r){
				if (c.getCellType() == Cell.CELL_TYPE_STRING) {
					
					String value = c.getStringCellValue();
					System.out.println(value);
				}
				else {
					if(c.getNumericCellValue()%1 == 0) {
						int value = (int) c.getNumericCellValue();
						System.out.println(value);
					}
					else {
						double value = c.getNumericCellValue();
						System.out.println(value);
					}
				}
				
			}
			fis.close();
			wb.close();
			
		}

	}

}
