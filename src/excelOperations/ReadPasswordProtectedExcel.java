package excelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadPasswordProtectedExcel {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream(".\\dataFiles\\customers.xlsx");
		String password = "test123";
//		XSSFWorkbook workbook = new XSSFWorkbook(fis);
//		Workbook workbook = WorkbookFactory.create(fis, password);
		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis, password);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows = sheet.getLastRowNum();
		int columns = sheet.getRow(0).getLastCellNum();
		
		System.out.println(rows);	// 5 (started from 0)
		System.out.println(columns);	// 3 (started from 1)
		
		/*
		//Read Data from sheet using for loop
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
				case FORMULA:
					System.out.print(cell.getNumericCellValue());
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		*/
		
		//Read Data from sheet using iterator
		
		Iterator<Row> iterator = sheet.iterator();
		while(iterator.hasNext())
		{
			Row nextrow = iterator.next();
			Iterator<Cell> celliterator = nextrow.cellIterator();
			while(celliterator.hasNext())
			{
				Cell cell = celliterator.next();
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
				case FORMULA:
					System.out.print(cell.getNumericCellValue());
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		
		workbook.close();

	}

}
