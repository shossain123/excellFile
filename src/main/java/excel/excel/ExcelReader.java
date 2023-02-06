package excel.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	String filename;
	String sheetName;
	Sheet sh;
	FileInputStream fis;
	Workbook wb;

	public ExcelReader(String f, String st) throws IOException {
		filename = f;
		sheetName = st;
		fis = new FileInputStream(f);
		wb = new XSSFWorkbook(fis);
		sh = wb.getSheet(st);
	}

	public Object[][] excelToArray() throws IOException {
		Object[][] table;

		fis = new FileInputStream(filename);
		wb = new XSSFWorkbook(fis);
		sh = wb.getSheet(sheetName);

		int Rows = sh.getPhysicalNumberOfRows();
		int Cols = sh.getRow(0).getPhysicalNumberOfCells();
		table = new Object[Rows - 1][Cols];

		for (int r = 1; r < Rows; r = r + 1) {
			for (int c = 0; c < Cols; c = c + 1) {
				Cell cell = sh.getRow(r).getCell(c);
				if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					String value = cell.getStringCellValue();
					table[r - 1][c] = value;

				} else {
					if (cell.getNumericCellValue() % 1 == 0) {
						int v = (int) cell.getNumericCellValue();
						table[r - 1][c] = "" + v;
					}

					else {
						double d = cell.getNumericCellValue();
						table[r - 1][c] = "" + d;
					}
				}
			}
		}

		return table;
	}

	public void updateCell(int row, int cell) throws IOException {
		// read the cell and update the cell and close the files.

		Cell c = sh.getRow(row).getCell(cell);
		fis.close();
		// wb.close();

		FileOutputStream fos = new FileOutputStream(filename);
		int value = Integer.parseInt(c.getStringCellValue().split(" ")[1]);

		c.setCellValue(c.getStringCellValue().split(" ")[0] + " " + (++value));
		wb.write(fos);
		fos.close();

	}

	public String getCellData(int row, int cell) throws IOException {
		String result = "";
		Cell c = sh.getRow(row).getCell(cell);

		if (c.getCellType() == Cell.CELL_TYPE_STRING) {
			result = c.getStringCellValue();

		} else {
			if (c.getNumericCellValue() % 1 == 0) {
				int v = (int) c.getNumericCellValue();
				result = "" + v;
			}

			else {
				double d = c.getNumericCellValue();
				result = "" + d;
			}


		}
		return result;

	}

}
