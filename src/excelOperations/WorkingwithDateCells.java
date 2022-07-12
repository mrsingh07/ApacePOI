package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkingwithDateCells {

	public static void main(String[] args) throws IOException {

		// Create black workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Date formats");

		// Date in normal format
		XSSFCell cell = sheet.createRow(0).createCell(0);
		cell.setCellValue(new Date()); // Current Date

		XSSFCreationHelper creationHelper = workbook.getCreationHelper();

		// format1: dd-mm-yyyy
		CellStyle style1 = workbook.createCellStyle();
		style1.setDataFormat(creationHelper.createDataFormat().getFormat("dd-mm-yyyy"));

		XSSFCell cell1 = sheet.createRow(1).createCell(0);
		cell1.setCellValue(new Date()); // Current Date
		cell1.setCellStyle(style1);

		// format2: mm-dd-yyyy
		CellStyle style2 = workbook.createCellStyle();
		style2.setDataFormat(creationHelper.createDataFormat().getFormat("mm-dd-yyyy"));

		XSSFCell cell2 = sheet.createRow(2).createCell(0);
		cell2.setCellValue(new Date()); // Current Date
		cell2.setCellStyle(style2);

		// format3: mm-dd-yyyy hh:mm:ss
		CellStyle style3 = workbook.createCellStyle();
		style3.setDataFormat(creationHelper.createDataFormat().getFormat("mm-dd-yyyy hh:mm:ss"));

		XSSFCell cell3 = sheet.createRow(3).createCell(0);
		cell3.setCellValue(new Date()); // Current Date
		cell3.setCellStyle(style3);

		// format4: hh:mm:ss
		CellStyle style4 = workbook.createCellStyle();
		style4.setDataFormat(creationHelper.createDataFormat().getFormat("hh:mm:ss"));

		XSSFCell cell4 = sheet.createRow(4).createCell(0);
		cell4.setCellValue(new Date()); // Current Date
		cell4.setCellStyle(style4);

		FileOutputStream fos = new FileOutputStream(".\\dataFiles\\dataformats.xlsx");
		workbook.write(fos);
		workbook.close();
		fos.close();

		System.out.println("Done!!!");

	}

}
