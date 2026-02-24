package SeleniumPractice;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriting {
	public static void main(String[] args) throws IOException {
		
		FileOutputStream excelfile = new FileOutputStream(System.getProperty("user.dir")+"\\Files\\Personal-data-writing.xlsx");
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("data");
		
		XSSFRow row1 = sheet.createRow(0);
		row1.createCell(0).setCellValue("name");
		row1.createCell(1).setCellValue("age");
		row1.createCell(2).setCellValue("education");
		row1.createCell(3).setCellValue("occupation");
		
		XSSFRow row2 = sheet.createRow(1);
		row2.createCell(0).setCellValue("priyanka");
		row2.createCell(1).setCellValue("31");
		row2.createCell(2).setCellValue("btech");
		row2.createCell(3).setCellValue("engineer");
		
		XSSFRow row3 = sheet.createRow(2);
		row3.createCell(0).setCellValue("Srinaya");
		row3.createCell(1).setCellValue("3");
		row3.createCell(2).setCellValue("nursery");
		row3.createCell(3).setCellValue("student");

		XSSFRow row4 = sheet.createRow(3);
		row4.createCell(0).setCellValue("rahul");
		row4.createCell(1).setCellValue("35");
		row4.createCell(2).setCellValue("MTech");
		row4.createCell(3).setCellValue("engineer");
		
		wb.write(excelfile);
		wb.close();
		excelfile.close();
		
		System.out.println("data writing done successfully");
		
	}

}
