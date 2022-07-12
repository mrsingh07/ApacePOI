package excelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Workbook --> Sheet --> Rows --> Cells
public class WritingExcelDemo1 {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		
		// 2D Array containing Heterogeneous Data
		Object empdata[][] = {	{"EmpID","Name","Job"},
								{101,"David","Engineer"},
								{102,"Smith","Manager"},
								{103,"Scott","Analyst"},
				
							};
		/*
		//Using for loop
		int rows = empdata.length;
		int columns = empdata[0].length;
		
		System.out.println(rows);	//4
		System.out.println(columns);	//3
		
		for(int r=0; r<rows; r++)
		{
			XSSFRow row = sheet.createRow(r);
			for(int c=0; c<columns; c++)
			{
				XSSFCell cell = row.createCell(c);
				Object value = empdata[r][c];
				
				if(value instanceof String)
				{
					cell.setCellValue((String)value);
				}
				else if(value instanceof Integer)
				{
					cell.setCellValue((Integer)value);
				}
				else if(value instanceof Boolean)
				{
					cell.setCellValue((Boolean)value);
				}
			}
		}
		*/
		
		// Using for each loop
		int rowCount = 0;
		for(Object emp[] : empdata)
		{
			XSSFRow row = sheet.createRow(rowCount++);
			int columnCount=0;
			for(Object value:emp)
			{
				XSSFCell cell = row.createCell(columnCount++);
				if(value instanceof String)
				{
					cell.setCellValue((String)value);
				}
				else if(value instanceof Integer)
				{
					cell.setCellValue((Integer)value);
				}
				else if(value instanceof Boolean)
				{
					cell.setCellValue((Boolean)value);
				}
			}
		}
		
		String filePath = ".\\dataFiles\\employee.xlsx";
		FileOutputStream outStream = new FileOutputStream(filePath);
		workbook.write(outStream);
		outStream.close();
		
		System.out.println("Employee.xlsx file written successfully...");
	}

}
