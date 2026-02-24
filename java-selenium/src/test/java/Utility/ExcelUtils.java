package Utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {

	public static FileInputStream fileIn;
	public static FileOutputStream fileOut;
	public static XSSFWorkbook wb;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static CellStyle style;

	public static int getRowCount(String filePath, String SheetName) throws IOException {
		fileIn = new FileInputStream(filePath);
		wb = new XSSFWorkbook(fileIn);
		sheet = wb.getSheet(SheetName);
		int rowCount = sheet.getLastRowNum();
		wb.close();
		fileIn.close();
		return rowCount;
	}

	public static int getCellCount(String filePath, String SheetName) throws IOException {
		fileIn = new FileInputStream(filePath);
		wb = new XSSFWorkbook(fileIn);
		sheet = wb.getSheet(SheetName);
		int cellCount = sheet.getRow(0).getLastCellNum();
		wb.close();
		fileIn.close();
		return cellCount;
	}

	public static String getCellData(String filePath, String SheetName, int rowNum, int colNum) throws IOException {
		fileIn = new FileInputStream(filePath);
		wb = new XSSFWorkbook(fileIn);
		sheet = wb.getSheet(SheetName);
		cell = sheet.getRow(rowNum).getCell(colNum);
		String data;

		try {
			// data = cell.toString();

			// Instead of toString, Apache POI provides DataFormatter class to convert the
			// all cell data into String

			DataFormatter format = new DataFormatter();
			data = format.formatCellValue(cell);

			// we can use either toString or DataFormatter. Both works fine
		} catch (Exception e) {
			data = "";
		}

		wb.close();
		fileIn.close();
		return data;
	}

	public static void setCelldata(String filePath, String SheetName, int rowNum, int colNum, String data) throws IOException {
		fileIn = new FileInputStream(filePath);
		wb = new XSSFWorkbook(fileIn);
		sheet = wb.getSheet(SheetName);
		row = sheet.getRow(rowNum);
		cell = row.createCell(colNum);
		cell.setCellValue(data);

		fileOut = new FileOutputStream(filePath);
		wb.write(fileOut);

		wb.close();
		fileIn.close();
		fileOut.close();
	}

	public static void fillGreenColour(String filePath, String SheetName, int rowNum, int colNum) throws IOException {

		fileIn = new FileInputStream(filePath);
		wb = new XSSFWorkbook(fileIn);
		sheet = wb.getSheet(SheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum);

		style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cell.setCellStyle(style);
		
		fileOut = new FileOutputStream(filePath);
		wb.write(fileOut);

		wb.close();
		fileIn.close();
		fileOut.close();

	}
	
	public static void fillRedColour(String filePath, String SheetName, int rowNum, int colNum) throws IOException {

		fileIn = new FileInputStream(filePath);
		wb = new XSSFWorkbook(fileIn);
		sheet = wb.getSheet(SheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum);

		style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cell.setCellStyle(style);
		
		fileOut = new FileOutputStream(filePath);
		wb.write(fileOut);

		wb.close();
		fileIn.close();
		fileOut.close();

	}
}
