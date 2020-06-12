package SkinnyPOM;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;


public class SkinnyPOM {
	
	WebDriver driver;
	 
    @BeforeMethod
	public void initialization() throws IOException{
	    
    	// To set the path of the Chrome driver.
		
    	System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		driver =new ChromeDriver();
		
		//To delete Cookies
		driver.manage().deleteAllCookies();
		
	    // To launch URL
	    driver.get("https://skinnyties.com/");
	    
	    // To maximize the browser
	    driver.manage().window().maximize();
	    
	    // implicit wait
	    driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
	    
	    
					
    }
		  
	@Test
	public void TC_01_IncQuality() throws IOException, InterruptedException{
		
		// Load the properties File		
	    Properties obj = new Properties();					
	    FileInputStream objfile = new FileInputStream(System.getProperty("user.dir")+"\\src\\SkinnyPOM\\application.properties");									
	    obj.load(objfile);
	    
	    // To Click on the Pattern Link
		driver.findElement(By.xpath(obj.getProperty("Pattern"))).click();
		
		// To Click on the Collection Link
		driver.findElement(By.xpath(obj.getProperty("Collection"))).click();
		
		// To Click on the Selected Tie Link
		driver.findElement(By.xpath(obj.getProperty("TieSelection"))).click();
		
		// To Click on the increasing the quality
		driver.findElement(By.xpath(obj.getProperty("Inc_Quality"))).click();
		
		// To Click on the CheckOut Link
		driver.findElement(By.xpath(obj.getProperty("Checkout"))).click();
		
		// To Click on the Close Button
		driver.findElement(By.xpath(obj.getProperty("CloseButton"))).click();
		
		// To get the value of the CART
		String ActualIncQuality = driver.findElement(By.xpath(obj.getProperty("CartValue"))).getText();
		
		String ExpectedIncQuality = "CART (2)";
		
		if (ExpectedIncQuality.equals(ActualIncQuality))
		{
			System.out.println("Number of items in the basket is increased by 1");
		}	
		driver.close();
		
		//To Print Status into Excel
		WriteintoExcel(1);
		
		
	}
	
	
	@Test
	public void TC_02_DecQuality() throws IOException, InterruptedException{
		
		// Load the properties File		
	    Properties obj = new Properties();					
	    FileInputStream objfile = new FileInputStream(System.getProperty("user.dir")+"\\src\\SkinnyPOM\\application.properties");									
	    obj.load(objfile);
	    
	    
	    // To Click on the Pattern Link
		driver.findElement(By.xpath(obj.getProperty("Pattern"))).click();
		
		// To Click on the Collection Link
		driver.findElement(By.xpath(obj.getProperty("Collection"))).click();
		
		// To Click on the Selected Tie Link
		driver.findElement(By.xpath(obj.getProperty("TieSelection"))).click();
		
		// To Click on the increasing the quality
		driver.findElement(By.xpath(obj.getProperty("Inc_Quality"))).click();
		
		// To Click on the increasing the quality
		driver.findElement(By.xpath(obj.getProperty("Inc_Quality"))).click();
		
		// To Click on the decreasing the quality
		driver.findElement(By.xpath(obj.getProperty("Dec_Quality"))).click();		
		
		// To Click on the CheckOut Link
		driver.findElement(By.xpath(obj.getProperty("Checkout"))).click();
		
		// To Click on the Close Button
		driver.findElement(By.xpath(obj.getProperty("CloseButton"))).click();
		
		// To get the value of the CART	
		String ActualDecQuality = driver.findElement(By.xpath(obj.getProperty("CartValue"))).getText();
		
		String ExpectedDecQuality = "CART (2)";
		
		if (ExpectedDecQuality.equals(ActualDecQuality))
		{
			System.out.println("Number of items in the basket is Decresed by 1");
		}
	
		driver.close();
		
		//To Print Status into Excel
		WriteintoExcel(2);
		
	}
	


public static void WriteintoExcel(int TC) throws IOException{
	
	// Import excel sheet.
	File src=new File(System.getProperty("user.dir")+"\\src\\SkinnyPOM\\ResultsExcel.xlsx");		  
	// Load the file.
	FileInputStream fis = new FileInputStream(src);
	// Load the workbook.
	Workbook workbook = new XSSFWorkbook(fis);
	
	// Load the sheet in which data is stored.
	Sheet sheet= workbook.getSheetAt(0);
	
	//To write data in the excel
	 FileOutputStream fos=new FileOutputStream(src);
	 
	 // Message to be written in the excel sheet
	     String message = "Pass";
	     
	     // Create cell where data needs to be written.
	     sheet.getRow(TC).createCell(2).setCellValue(message);
	         
	     // finally write content
	     workbook.write(fos);

	 // close the file
	 fos.close();	
	
}
}
