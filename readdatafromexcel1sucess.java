package TestNGDemo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;



public class ReadExcel {
	
	@Test
	public void readdataFromExcel() throws IOException {
		FileInputStream file = new FileInputStream("F:\\software testing and Automation\\Recording Lecture\\sql part\\SQL SCREEN SHOT\\SQL DATA BASE.Xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("DETAIL TABLE 1 (2)");
		
		System.out.println(sheet.getRow(4).getCell(4).getStringCellValue());
		System.out.println(sheet.getRow(1).getCell(0).getNumericCellValue());
		
		// writing data in excel sheet
		Row row = sheet.createRow(14);
		Cell cell = row.createCell(5);
		cell.setCellValue("Utkarshaa Academy Pune");
		FileOutputStream fos = new FileOutputStream("F:\\software testing and Automation\\Recording Lecture\\sql part\\SQL SCREEN SHOT\\SQL DATA BASE.Xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("writing data in excel sucessfully");
	}

}
