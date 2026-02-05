package SeleniumPractice;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WindowType;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FileUpload {
	
	public static void main(String[] args) throws InterruptedException {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5000));
		
		driver.get("https://davidwalsh.name/demo/multiple-file-upload.php");
		
		//single file uploading
		
		String fileName1 = "E:\\new-automation\\my-java-selenium-practice\\java-selenium\\Files\\Name.txt";
		driver.findElement(By.id("filesToUpload")).sendKeys(fileName1);
		if(driver.findElement(By.cssSelector("#fileList >li")).getText().equals("Name.txt")) {
			System.out.println("File uploaded successfully");
		}else {
			System.out.println("Failed to upload the file");
		}
		
		Thread.sleep(2000);
		
		driver.switchTo().newWindow(WindowType.TAB);
		driver.get("https://davidwalsh.name/demo/multiple-file-upload.php");
		//Multiple file uploading
		String fileName2 = "E:\\new-automation\\my-java-selenium-practice\\java-selenium\\Files\\Occupation.txt";
		driver.findElement(By.id("filesToUpload")).sendKeys(fileName1+"\n"+fileName2);
		
		if(driver.findElement(By.xpath("//ul[@id='fileList']/li[1]")).getText().equals("Name.txt") && driver.findElement(By.xpath("//ul[@id='fileList']/li[2]")).getText().equals("Occupation.txt")) {
			System.out.println("Files uploaded successfully");
		}else {
			System.out.println("Failed to upload the file");
		}
		Thread.sleep(2000);
		
		driver.quit();
	}

}
