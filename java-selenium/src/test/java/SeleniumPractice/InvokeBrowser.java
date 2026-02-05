package SeleniumPractice;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class InvokeBrowser {
	
	public static void main(String[] args) {
		
		//selenium 3 required to mention the driver files explicitely
		//System.setProperty("webdriver.chrome.driver", "./drivers/chromedriver.exe");
		
		//above is deprecated in selenium4 and we need to use webdrivermanager
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.manage().timeouts().scriptTimeout(Duration.ofSeconds(10));
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(10));
		
		driver.get("https://www.youtube.com/");
		driver.findElement(By.name("search_query")).sendKeys("priyanka");
		
		driver.quit();
	}

}
