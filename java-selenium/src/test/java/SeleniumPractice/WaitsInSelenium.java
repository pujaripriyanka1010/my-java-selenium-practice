package SeleniumPractice;

import java.time.Duration;
import java.util.NoSuchElementException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class WaitsInSelenium {

	public static void main(String[] args) throws InterruptedException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();

		driver.manage().window().maximize();

		// 1st implicite wait
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.get("https://testautomationpractice.blogspot.com/");
		
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(5));

		// 2nd explicit wait
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.id("name"))));
		driver.findElement(By.id("name")).sendKeys("priyanka");

		Thread.sleep(2000);

		// 3rd fluent wait
		Wait<WebDriver> fluentWait = new FluentWait<WebDriver>(driver).pollingEvery(Duration.ofNanos(200))
				.ignoring(NoSuchElementException.class).withTimeout(Duration.ofSeconds(5));
		fluentWait.until(ExpectedConditions.elementToBeClickable(By.id("female")));
		
		driver.findElement(By.id("female")).click();
		Thread.sleep(2000);
		
		driver.quit();

	}
}
