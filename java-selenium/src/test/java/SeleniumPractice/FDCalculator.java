package SeleniumPractice;

import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import Utility.ExcelUtils;
import io.github.bonigarcia.wdm.WebDriverManager;

public class FDCalculator {
	
	public static void main(String[] args) throws IOException {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.manage().window().maximize();
		
		driver.get("https://www.sbisecurities.in/calculators/fd-calculator");
		
		String filePath = System.getProperty("user.dir")+"\\Files\\Personal-data.xlsx";
		ExcelUtils excel = new ExcelUtils();
		int totalRow = excel.getRowCount(filePath, "FD-details");
		int totalCol = excel.getCellCount(filePath, "FD-details");
		
		for(int i = 1;i<=totalRow;i++) {
			
			//Read data from Excelsheet
			String principle = excel.getCellData(filePath, "FD-details", i, 0);
			String years = excel.getCellData(filePath, "FD-details", i, 1);
			String ROI = excel.getCellData(filePath, "FD-details", i, 2);
			String freq = excel.getCellData(filePath, "FD-details", i, 3);
			String expAmount = excel.getCellData(filePath, "FD-details", i, 4);
			
			//Perform operation using above data
			driver.findElement(By.name("fd_investment")).clear();
			driver.findElement(By.name("fd_investment")).sendKeys(principle);
			
			driver.findElement(By.name("years")).clear();
			driver.findElement(By.name("years")).sendKeys(years);
			
			driver.findElement(By.id("input_interest")).clear();
			driver.findElement(By.id("input_interest")).sendKeys(ROI);
			
			
			Select freqDD = new Select(driver.findElement(By.cssSelector("#frequency")));
			freqDD.selectByVisibleText(freq);
			
			driver.findElement(By.xpath("//button[@type='submit']")).click();
			
			String actualAmount = driver.findElement(By.xpath("//div[@id='lumpsum_return']/p[1]/span")).getText().replace(",", "").replace("₹", "");
			
			if(Double.parseDouble(actualAmount)==Double.parseDouble(expAmount)) {
				System.out.println("test pass");
				excel.setCelldata(filePath, "FD-details", i, 5, actualAmount);
				excel.setCelldata(filePath, "FD-details", i, 6, "PASS");
				excel.fillGreenColour(filePath, "FD-details", i, 6);
			}else {
				System.out.println("test fail");
				excel.setCelldata(filePath, "FD-details", i, 5, actualAmount);
				excel.setCelldata(filePath, "FD-details", i, 6, "FAIL");
				excel.fillRedColour(filePath, "FD-details", i, 6);
			}
			
		}
		
		driver.quit();
		
		
	}

}
