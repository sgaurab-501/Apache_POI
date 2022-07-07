package com.qa.dataDrivenTesting;

import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.qa.util.ExcelUtil;

public class WebTable2Excel {

public static void main(String[] args) throws Exception {
	
	System.setProperty("webdriver.chrome.driver", ".\\Drivers\\chromedriver_win32\\chromedriver.exe");
WebDriver driver = new ChromeDriver();

WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

WebElement footerLink = driver.findElement(By.xpath("//li[@id='footer-info-lastmod']"));
wait.until(ExpectedConditions.visibilityOf(footerLink));

driver.manage().window().maximize();
driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");

String excelFilePath = ".\\DataFiles\\Population.xlsx";

ExcelUtil eUtil = new ExcelUtil(excelFilePath);

// Write Header in excel sheet

eUtil.setCellData("Sheet1", 0, 0, "Country");
eUtil.setCellData("Sheet1", 0, 1, "Region");
eUtil.setCellData("Sheet1", 0, 2, "Population");
eUtil.setCellData("Sheet1", 0, 3, "Percentage of the World");
eUtil.setCellData("Sheet1", 0, 4, "Date");

WebElement table= driver.findElement(By.xpath("//table[@class='wikitable sortable jquery-tablesorter']/tbody"));

int rows = table.findElement(By.xpath("tr")).getSize(); // no. of rows present in web table


for(int r=1; r<=rows; r++) {
	
//	driver.findElement(By.xpath("//table[@class='wikitable sortable jquery-tablesorter']/tbody/tr/td[r]"));

	table.findElement("tr[1]/td[r] ");

}





	
	
}	
	
}
