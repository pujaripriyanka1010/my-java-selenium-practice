package SeleniumPractice;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FrameHandling {
	
	public static void main(String[] args) {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.manage().timeouts().scriptTimeout(Duration.ofSeconds(10));
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(10));
		
		driver.get("https://bonigarcia.dev/selenium-webdriver-java/iframes.html");
		
		int size = driver.findElements(By.tagName("iframe")).size();
		System.out.println("total iframes are => "+size);
		
		driver.switchTo().frame("my-iframe");
		System.out.println(driver.getTitle());
		
		
		
		driver.switchTo().defaultContent();
		System.out.println(driver.getTitle());
		
		driver.quit();
	}

}
