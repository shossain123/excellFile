package excel.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BankRead {

	public static void main(String[] args) throws IOException {
		 
		String filename = "data/Bank.xlsx";
		FileInputStream fis = new FileInputStream(filename);
		Workbook wb = new XSSFWorkbook(fis);
	    Sheet sh =	wb.getSheet("Sheet1");
	    int totalRow = sh.getPhysicalNumberOfRows();
	    int totalCell = sh.getRow(0).getPhysicalNumberOfCells();
	    
	    for(int r=1; r<totalRow ; r= r+1) {
	    	for(int c=0; c<totalCell ; c=c+1) {
	    		
	    		Cell cell = sh.getRow(r).getCell(c);
	    		if(cell.getCellType()==Cell.CELL_TYPE_STRING) {
	    			 String s =    cell.getStringCellValue();
	    			 System.out.println(s);
	    		}
	    		else{
	    			if(cell.getNumericCellValue()%1==0) {
	    			int value = (int) cell.getNumericCellValue();
	    			System.out.println(value);
	    			}
	    			else {
	    			double d =	cell.getNumericCellValue();
	    			System.out.println(d);
	    			}
	    		}
	    	}
	    	
	    }
	

	}

}
