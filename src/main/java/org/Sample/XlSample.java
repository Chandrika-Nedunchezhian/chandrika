package org.Sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class XlSample {

	public static void main(String[] args) throws IOException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.amazon.in/");
		
		WebElement btnSearch = driver.findElement(By.id("twotabsearchtextbox"));
		btnSearch.sendKeys("hp laptops",Keys.ENTER);
		List<WebElement> hpLaptop = driver.findElements(By.xpath("//span[@class='a-size-base-plus a-color-base a-text-normal']"));
		File f=new File("C:\\Users\\User\\eclipse-chandrika\\MavenBrowserLaunch\\XlReader\\Book1.xlsx");
		FileInputStream stream=new FileInputStream(f);
	    Workbook workBook=new XSSFWorkbook(stream);
	    Sheet createSheet = workBook.createSheet("Hp Laptop");
	    for(int i=0;i<hpLaptop.size();i++) {
	    	Row row = createSheet.createRow(i);
	    	Cell cell = row.createCell(0);
	    	WebElement e = hpLaptop.get(i);
	    	String text = e.getText();
	    	cell.setCellValue(text);
	    	FileOutputStream fos=new FileOutputStream(f);
	    	workBook.write(fos);
	    	
	    	
	    }
}
}