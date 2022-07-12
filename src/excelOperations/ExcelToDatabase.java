package excelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDatabase {

	public static void main(String[] args) throws SQLException, IOException {
		
		//Database Connection
		Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/hr", "root", "Harry@123");
		Statement stmt = con.createStatement();
		
		//Create a new table in the database 'places'
		String sql = "create table places (LOCATION_ID decimal(4,0), STREET_ADDRESS varchar(40), POSTAL_CODE varchar(12), CITY varchar(30), STATE_PROVINCE varchar(25), COUNTRY_ID varchar(2));";
		stmt.execute(sql);
		
		FileInputStream fis = new FileInputStream(".\\dataFiles\\locations.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows = sheet.getLastRowNum();
		
		for(int r=1; r<=rows; r++)
		{
			XSSFRow row = sheet.getRow(r);
			double locId = row.getCell(0).getNumericCellValue();
			String streetAdd = row.getCell(1).getStringCellValue();
			String postalCode = row.getCell(2).getStringCellValue();
			String city = row.getCell(3).getStringCellValue();
			String stateProvince = row.getCell(4).getStringCellValue();
			String conId = row.getCell(5).getStringCellValue();
			
			String insertion = "insert into places values('"+locId+"','"+streetAdd+"','"+postalCode+"','"+city+"','"+stateProvince+"','"+conId+"');";
			stmt.execute(insertion);
			stmt.execute("commit");
		}
		workbook.close();
		fis.close();
		con.close();
		
		System.out.println("Done!!!");
		

	}

}
