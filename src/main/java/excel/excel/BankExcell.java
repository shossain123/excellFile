package excel.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BankExcell {
         String filename;
         String sheetname;
         Sheet sh;
         Workbook wb;
         FileInputStream fis;
      public BankExcell(String f, String s) throws IOException {
    	  filename = f;
    	  sheetname = s;
    	  fis = new FileInputStream(f);
    	   wb = new XSSFWorkbook(fis);
    	     sh = wb.getSheet(s);
    	     
      }
      
      public Object[][] excellToArray(){
    	  Object[][] table;
    	  int Rows = sh.getPhysicalNumberOfRows();
    	  int Cols = sh.getRow(0).getPhysicalNumberOfCells();
    	  table = new Object[Rows-1][Cols];
    	  for(int r=1; r<Rows; r=r+1) {
    		  for(int c=0; c<Cols; c=c+1) {
    			  Cell cell= sh.getRow(r).getCell(c);
    			  if(cell.getCellType()== Cell.CELL_TYPE_STRING) {
    				String value= cell.getStringCellValue();
    				table[r-1][c]=value;
    				}
    			  else {
    				  if(cell.getNumericCellValue()%1==0) {
    				int value =	(int) cell.getNumericCellValue();
    				table[r-1][c]= ""+value;
    				  }
    				  else {
    					  double d = cell.getNumericCellValue();
    					  table[r-1][c] = ""+d;
    				  }
    				  
    			  }
    			  
    		  }
    	  }
    	  
    	  
    	  
    	  return table;
      }
	
}
