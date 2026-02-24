package SeleniumPractice;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReading {
	public static void main(String[] args) throws IOException {
		
		FileInputStream excelFile = new FileInputStream(System.getProperty("user.dir")+"\\Files\\Personal-data.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(excelFile);
		XSSFSheet sheet = wb.getSheetAt(0);
		int totalRow = sheet.getLastRowNum();// index start from 0
		int totalclm = sheet.getRow(totalRow).getLastCellNum(); // index start from 1
		System.out.println("total row and columns are : "+totalRow+" & "+totalclm);
		
		for (int i = 0; i <= totalRow; i++) {
			XSSFRow row = sheet.getRow(i);
			for (int j = 0; j < totalclm; j++) {
				String cellData = row.getCell(j).toString();
				System.out.print(cellData +"\t");
			}
			System.out.println();
		}
	
		
		wb.close();
		excelFile.close();
	
	
	
	}

}
