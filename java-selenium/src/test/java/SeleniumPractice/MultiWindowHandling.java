package SeleniumPractice;

import java.time.Duration;
import java.util.Set;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WindowType;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class MultiWindowHandling {
	
	public static void main(String[] args) throws InterruptedException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(5));
		
		driver.get("https://www.youtube.com/");
		
		String parentHandle = driver.getWindowHandle();
		System.out.println("parent handle => "+parentHandle);
		
		for(int i = 0;i<5;i++) {
			driver.switchTo().newWindow(WindowType.TAB);
		}
		
		Set<String> handles = driver.getWindowHandles();
		
		System.out.println(handles);
		
		Thread.sleep(2000);
		driver.get("https://www.google.com/");
		driver.close();
		
		driver.switchTo().window(parentHandle);		
		
		Thread.sleep(2000);
		System.out.println(driver.getTitle());
		driver.quit();
		
	}

}
