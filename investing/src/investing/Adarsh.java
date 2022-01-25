package investing;

import java.io.File;
import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementClickInterceptedException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Adarsh {

	public static void main(String[] args) throws Exception {
		File abpath=new File("./resources/Name.xlsx");
		FileInputStream fis=new FileInputStream(abpath);
		Workbook workbook = WorkbookFactory.create(fis);
		Sheet sheet = workbook.getSheet("Sheet1");
		System.setProperty("webdriver.chrome.driver", "./drivers/chromedriver.exe");
		WebDriver driver= new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.investing.com/");
		JavascriptExecutor js=(JavascriptExecutor) driver;
		try {
			WebElement close = driver.findElement(By.xpath("//i[@class='popupCloseIcon largeBannerCloser']"));
			js.executeScript("arguments[0].click()",close);
			
		} catch (NoSuchElementException e) {
			// TODO: handle exception
		
		}
		driver.findElement(By.linkText("No thanks")).click();
		WebElement signinlink = driver.findElement(By.linkText("Sign In"));
		signinlink.click();
		WebElement emailtextfield = driver.findElement(By.id("loginFormUser_email"));
		WebElement passwordtextfield = driver.findElement(By.id("loginForm_password"));
		WebElement Signinbutton = driver.findElement(By.xpath("//*[@id=\"signup\"]/a"));
		emailtextfield.clear();
		emailtextfield.sendKeys("adarshangadi13@gmail.com");
		passwordtextfield.clear();
		passwordtextfield.sendKeys("Adarsh@13");
		Signinbutton.click();
		WebDriverWait wait=new WebDriverWait(driver, 10);
		for(int i=1;i<=20;i++) {
		WebElement search = driver.findElement(By.xpath("//input[contains(@placeholder,'Search the website')]"));
		String data = sheet.getRow(i).getCell(0).getStringCellValue();
		search.sendKeys(data);
		WebElement name = driver.findElement(By.xpath("//span[contains(text(),'"+data+"')]"));
		wait.until(ExpectedConditions.visibilityOf(name));
		name.click();
		try {
			driver.findElement(By.xpath("//span[contains(text(),'"+data+"')]")).click();
		} catch (NoSuchElementException e) {
			// TODO: handle exception
			
		}
		for(;;) {
			try { 
				driver.findElement(By.xpath("//*[@id=\"__next\"]/div/div/div[2]/main/div/div[6]/nav/ul/li[3]/a")).click();
				break;
			}catch(NoSuchElementException e) {
				js.executeScript("window.scrollBy(0,550)");
			}
			catch (ElementClickInterceptedException e) {
				// TODO: handle exception
				driver.findElement(By.id("closeIconHit")).click();
			}
		}
		js.executeScript("window.scrollBy(0,550)");
		WebElement datefield = driver.findElement(By.id("widgetFieldDateRange"));
		for(;;) {
			try { 
		wait.until(ExpectedConditions.elementToBeClickable(datefield));
				datefield.click();
				break;
			}catch(TimeoutException e) {
				js.executeScript("arguments[0].click()", driver.findElement(By.xpath("//*[@id=\"datePickerIconWrap\"]/span")));
			}
		}
		driver.findElement(By.id("startDate")).clear();
		driver.findElement(By.id("startDate")).sendKeys("01/01/2015");
		driver.findElement(By.id("applyBtn")).click();
		driver.findElement(By.xpath("//a[@title='Download Data']")).click();
		}
	}
	}
