package SeleniumPractice;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WindowType;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class JSExecutor {
	
	public static void main(String[] args) throws InterruptedException {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5000));
		
		driver.get("https://www.youtube.com/");
		
		//js for sendkeys
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].setAttribute('value','priyanka')", driver.findElement(By.name("search_query")));
		Thread.sleep(2000);
		
		//js for click method
		String parentID = driver.getWindowHandle();
		driver.switchTo().newWindow(WindowType.TAB);
		driver.get("https://www.worldometers.info/geography/flags-of-the-world/");
		
		js.executeScript("arguments[0].click()", driver.findElement(By.xpath("//*[@alt='Australia']")));
		Thread.sleep(2000);
		driver.switchTo().window(parentID);
		
		//js for scroll into view
		driver.switchTo().newWindow(WindowType.WINDOW);
		driver.get("https://www.worldometers.info/geography/flags-of-the-world/");
		
		js.executeScript("arguments[0].scrollIntoView()", driver.findElement(By.xpath("//*[@alt='India']")));
		
		
		Thread.sleep(2000);
		//js for scroll up and down
		
		js.executeScript("window.scrollBy(0,document.body.scrollHeight)");
		Thread.sleep(2000);
		
		js.executeScript("window.scrollBy(0,-document.body.scrollHeight)");
		Thread.sleep(2000);
		
		
		driver.quit();
	}

}
