package excel.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Excellread2 {

	public static void main(String[] args) throws IOException {
		String file = "data/sample1.xls";

		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new HSSFWorkbook(fis);
		Sheet sh = wb.getSheet("Sheet3");
		int totalRow = sh.getPhysicalNumberOfRows();
		int tcell = sh.getRow(0).getPhysicalNumberOfCells();
		System.out.println(totalRow);
		System.out.println(tcell);

		for (int i = 0; i < totalRow; i = i + 1) {

			for (int c1 = 0; c1 < tcell; c1++) {

				Cell c = sh.getRow(i).getCell(c1);
				if (c.getCellType() == Cell.CELL_TYPE_STRING) {

					String value = c.getStringCellValue();
					System.out.println(value);
				} else {
					if (c.getNumericCellValue() % 1 == 0) {
						int value = (int) c.getNumericCellValue();
						System.out.println(value);
					} else {
						double value = c.getNumericCellValue();
						System.out.println(value);
					}
				}

			}
		}
		wb.close();
		fis.close();

	}

}
