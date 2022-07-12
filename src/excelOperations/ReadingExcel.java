package excelOperations;

/*
 * Reading excel file
 */

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath = ".\\dataFiles\\countries.xlsx";
		
		FileInputStream inputStream = new FileInputStream(excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
//		XSSFSheet sheet = workbook.getSheet("Sheet1");
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		/*
		//Using for loop
		int rows = sheet.getLastRowNum();
		int columns = sheet.getRow(1).getLastCellNum();
		
		for(int r=0; r<=rows; r++)
		{
			XSSFRow row = sheet.getRow(r);
			for(int c=0; c<columns; c++)
			{
				XSSFCell cell = row.getCell(c);
				
				switch(cell.getCellType())
				{
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		*/
		
		// Using Iterator
		Iterator iterator = sheet.iterator();
		while(iterator.hasNext())
		{
			XSSFRow row = (XSSFRow) iterator.next();
			
			Iterator cellIterator = row.cellIterator();
			while(cellIterator.hasNext())
			{
				XSSFCell cell = (XSSFCell) cellIterator.next();
				
				switch(cell.getCellType())
				{
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				
				}
				System.out.print(" | ");
			}
			System.out.println();
		}

	}

}
