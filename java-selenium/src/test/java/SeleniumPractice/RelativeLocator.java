package SeleniumPractice;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import static org.openqa.selenium.support.locators.RelativeLocator.*;

import io.github.bonigarcia.wdm.WebDriverManager;

public class RelativeLocator {
	public static void main(String[] args) throws InterruptedException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();

		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(5));

		driver.get("https://testautomationpractice.blogspot.com/");

		// above()
		WebElement email = driver.findElement(By.cssSelector("#email"));
		driver.findElement(with(By.tagName("input")).above(email)).sendKeys("priyanka");
		Thread.sleep(2000);

		// below
		WebElement phone = driver.findElement(By.id("phone"));
		driver.findElement(with(By.tagName("textarea")).below(phone)).sendKeys("shrushti homes,pune");
		Thread.sleep(2000);

		// toLeftOf()
		WebElement femaleText = driver.findElement(By.cssSelector("[for='female']"));
		driver.findElement(with(By.tagName("input")).toLeftOf(femaleText)).click();
		Thread.sleep(2000);

		// near()
		WebElement femaletext = driver.findElement(By.cssSelector("[for='female']"));
		driver.findElement(with(By.id("tuesday")).toRightOf(femaletext)).click();
		Thread.sleep(2000);

		// toRightOf
		WebElement section1 = driver.findElement(By.id("input1"));
		driver.findElement(with(By.id("btn1")).toRightOf(section1)).click();
		Thread.sleep(2000);

		driver.quit();
	}

}
