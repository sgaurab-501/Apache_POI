package com.qa.dataDrivenTesting;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataDrivenTest {

	WebDriver driver;
	
	
@BeforeClass
public void setup() {
	
	System.setProperty("webdriver.chrome.driver", ".\\Drivers\\chromedriver_win32\\chromedriver.exe");
	driver = new ChromeDriver();
	
	WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	driver.manage().window().maximize();
	
}

@Test(dataProvider="LoginData")
public void loginTest(String uname, String pwd, String result) {

	driver.get("https://admin-demo.nopcommerce.com/login?ReturnUrl=%2Fadmin%2F");
WebElement emailField =	driver.findElement(By.id("Email"));
	emailField.clear();
	emailField.sendKeys(uname);
	
WebElement pwdField = driver.findElement(By.id("Password"));
pwdField.clear();
pwdField.sendKeys(pwd);

// System.out.println(uname+" "+pwd+" "+exp);

driver.findElement(By.xpath("//div[@class='buttons']/button[@type='submit']")).click(); //Login Button

String exp_Title = "Dashboard / nopCommerce administration";

String act_Title = driver.getTitle();
	
if(result.equalsIgnoreCase("Valid")) {
	
	if(exp_Title.equalsIgnoreCase(act_Title)) {
		
		Assert.assertTrue(true);

				}
	else {
		
	Assert.assertFalse(true);	
	}}

	else if(result.equalsIgnoreCase("Invalid")) {
		
		if(exp_Title.equalsIgnoreCase(act_Title)) {
			WebElement logoutLink =	driver.findElement(By.xpath("//li[@class='nav-item']/a[contains(text(),'Logout')]"));
			logoutLink.click();
			Assert.assertTrue(false);
		}
		
		else {
		Assert.assertTrue(true);	
		
	}
}
 

	
	
}


@DataProvider(name="LoginData")
	public Object[][] getData() {
	
	String loginData[][] = {{"admin@yourstore.com", "admin", "Valid"},
			{"admin@yourstore.com", "adm", "Invalid"},
			{"adm@yourstore.com", "admin", "Invalid"},
			{"adm@yourstore.com", "adm", "Invalid"}
	};
	
		return loginData;
}	
		

@AfterClass
public void tearDown() {
	
	driver.close();
}



	}
	
	
	
	

