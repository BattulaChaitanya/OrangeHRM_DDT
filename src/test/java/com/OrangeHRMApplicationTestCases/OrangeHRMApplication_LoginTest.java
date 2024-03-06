package com.OrangeHRMApplicationTestCases;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import com.BaseTests.BaseTestOrangeHRMApplication;
import com.Utility.Log;
import com.google.j2objc.annotations.Property;


public class OrangeHRMApplication_LoginTest extends BaseTestOrangeHRMApplication
{
	
	FileInputStream testDataFile;
	XSSFWorkbook workBook;
	XSSFSheet testDataSheet;
	
	@Test(priority = 1,description = "Adding Test Data into Excel")
	public void ExcelTestData() throws IOException 
	{
		testDataFile=new FileInputStream("./src/main/java/com/OrangeHRMApplicationTestDataFiles/OrangeHRMApplicationLoginTestExcelSheet.xlsx");
		workBook=new XSSFWorkbook(testDataFile);
		testDataSheet=workBook.getSheet("Sheet1");
		
		
		Row row1=testDataSheet.createRow(1);
		Cell userName=row1.createCell(7);
		userName.setCellValue("Admin");
		Cell password=row1.createCell(8);
		password.setCellValue("Chaitu@022");
		
		Row testDataRow2=testDataSheet.createRow(2);
		userName=testDataRow2.createCell(7);
		userName.setCellValue("Admin1");
		password=testDataRow2.createCell(8);
		password.setCellValue("Chaitu@022");
		
		
		Row testDataRow3=testDataSheet.createRow(3);
		userName=testDataRow3.createCell(7);
		userName.setCellValue("Admin");
		password=testDataRow3.createCell(8);
		password.setCellValue("Chaitanya");

		Row testDataRow4=testDataSheet.createRow(4);
		userName=testDataRow4.createCell(7);
		userName.setCellValue("Admin");
		password=testDataRow4.createCell(8);
		password.setCellValue("Chaitu@022");

		Row testDataRow5=testDataSheet.createRow(5);
		userName=testDataRow5.createCell(7);
		userName.setCellValue("Admin1");
		password=testDataRow5.createCell(8);
		password.setCellValue("Chaitanya");
		
		
		
	}
	
	@Test(priority = 2,description = "Writing Multiple Data in the Excel Sheet")
	public void Validating_LoginPageTitle() throws IOException
	{
		
		int rowCount=testDataSheet.getLastRowNum();
		for(int index=1;index<=rowCount;index++)
		{
			
			Row row=testDataSheet.getRow(index);
			
			Cell Expected_LoginPageTextOfCell0=row.createCell(0);
			Expected_LoginPageTextOfCell0.setCellValue("LOGIN Panel");
			
			Cell Expected_LoginPageTextOfCell1=row.createCell(1);
			Expected_LoginPageTextOfCell1.setCellValue("Invalid credentials");
			
			Cell Expected_LoginPageTitle=row.createCell(4);
			Expected_LoginPageTitle.setCellValue("orangeHRM");
			
			Cell Expected_HomePageText=row.createCell(9);
			Expected_HomePageText.setCellValue("Admin");
			
			
		}
		 
	}
	
	FileInputStream propertiesFile;
	Properties properties;
	
	
	
	@Test(priority = 3,description = "validating the LoginPage OrangeHRM LoGo")
	public void ValidatingLogo() throws IOException 
	{
		propertiesFile=new FileInputStream("./src/main/java/com/Config/OrangeHRMApplication_LoginTest.properties");
		properties=new Properties();
		properties.load(propertiesFile);
		
		By  loginPageLoGoimageProperty=By.xpath(properties.getProperty("orangeHRMApplicationLoginPageImageProperty"));
			WebElement loginPageLoGoimage = driver.findElement(loginPageLoGoimageProperty);
		 
		 boolean flag = loginPageLoGoimage.isDisplayed();
		 
		 if(flag)
		 {
			 Log.info("OrangeHRM Application Login Page OrangeHRM Image Uploaded Successfully - PASS");
		 }
		 else
		 {
			 Log.info("OrangeHRM Application Login Page OrangeHRM Image Failed to Upload - FAIL");
		 }
		
		
	}

	
	
	@Test(priority = 4,description = "OrangeHRM Application Login Functionality with Multiple Data")
	public void Validating_LoginTest() throws IOException
	{
		
		int rowCount=testDataSheet.getLastRowNum();
		for(int rowIndex=1;rowIndex<=rowCount;rowIndex++)
		{	
			Row row=testDataSheet.getRow(rowIndex);
			
			String expected_LoginPageTitle=row.getCell(4).getStringCellValue();
			
			Log.info("The Expected Title for the OrangeHRM LoginPage is :- "+expected_LoginPageTitle);
	    
	    	String actual_LoginPageTitle=driver.getTitle();
	    	Log.info("The Actual Title for the OrangeHRM LoginPage is :- "+actual_LoginPageTitle);
	    
	    	Cell Actual_LoginPageTitle=row.createCell(5);
	    	Actual_LoginPageTitle.setCellValue(actual_LoginPageTitle);
	    
	    	if(actual_LoginPageTitle.equals(expected_LoginPageTitle))
	    	{
	    		Log.info("Title for the OrangeHRM LoginPage is Matched - Pass");
	    			Cell TitleTestResult=row.createCell(6);
	    			TitleTestResult.setCellValue("Pass");
	        
	    	}
	    	else
	    	{
	    		Log.info("Title for the OrangeHRM LoginPage is Not Matched - Fail");
	    			Cell TitleTestResult=row.createCell(6);
	    			TitleTestResult.setCellValue("Fail");
	    	}

			
			String userNameTestData=row.getCell(7).getStringCellValue();
				By userNameProperty=By.id(properties.getProperty("orangeHRMApplicationLogInPageUserNameProperty"));
					WebElement userName=driver.findElement(userNameProperty);
						userName.sendKeys(userNameTestData);
			
			String passwordTestData=row.getCell(8).getStringCellValue();
			Actions keyboardAction=new Actions(driver);
			keyboardAction.sendKeys(Keys.TAB).build().perform();
			keyboardAction.sendKeys(passwordTestData).build().perform();
			
			keyboardAction.sendKeys(Keys.TAB).build().perform();
			keyboardAction.sendKeys(Keys.ENTER).build().perform();
			
		try 
		{
				//validating Home Page Text
				
				By adminProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageAdminProperty"));
					WebElement admin = driver.findElement(adminProperty);
						String adminText=admin.getText();
				
				Cell Actual_HomePageText=row.createCell(10);
				Actual_HomePageText.setCellValue(adminText);
				
				String expected_HomePageText=row.getCell(9).getStringCellValue();
				Log.info("The Expected Home Page Text is :- "+expected_HomePageText);
				
				String actual_HomePageText=row.getCell(10).getStringCellValue();
				Log.info("The Actual Home Page Text is:- "+actual_HomePageText);
				
				if(actual_HomePageText.equals(expected_HomePageText))
				{
					Log.info("The Actual Home Page Text is same as Expected Home Page Text - Pass");
						Cell TextTestResult=row.createCell(11);
						TextTestResult.setCellValue("Pass");
				}
				else
				{
					Log.info("The Actual Home Page Text is Not same as Expected Home Page Text - Fail");
						Cell TextTestResult=row.createCell(11);
						TextTestResult.setCellValue("Fail");
				}
				
				Cell Actual_LogInPageText=row.createCell(2);
				Actual_LogInPageText.setCellValue("Valid credentials");
				
				String expected_loginPageTextOfCell1=row.getCell(1).getStringCellValue();
				Log.info("The Expected span Message Text is:- "+expected_loginPageTextOfCell1);
				
				String actual_LoginPageText=row.getCell(2).getStringCellValue();
				Log.info("The Actual Span Message Text is:- "+actual_LoginPageText);
				
				if(Actual_LogInPageText.equals(expected_loginPageTextOfCell1))
				{
					Log.info("The Actual Span Message is same as Expected Span Message - Pass");
						Cell TextTestResult=row.createCell(3);
						TextTestResult.setCellValue("Pass");
				}
				else
				{
					Log.info("The Actual Span Message is Not same as Expected Span Message - Fail");
						Cell TextTestResult=row.createCell(3);
						TextTestResult.setCellValue("Fail");
				}
				
				
				By welcomeAdminproperty=By.id(properties.getProperty("orangeHRMApplicationHomePageWelcomeAdminProperty"));
					WebElement welcomeAdmin=driver.findElement(welcomeAdminproperty);
						welcomeAdmin.click();
				
			
				By logOutProperty=By.linkText(properties.getProperty("orangeHRMApplicationHomePageLogOutProperty"));	
					WebElement logOut=driver.findElement(logOutProperty);
					logOut.click();
				
		}
		catch(Exception e)
		{
		
			
				File CapturingScreenShort=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(CapturingScreenShort,new File("./ApplicationScreenShorts/OrangeHRMApplication UserName "+userNameTestData+" Password "+passwordTestData+".png"));
			
				
				By loginPanelProperty=By.id(properties.getProperty("orangeHRMApplicationLoginPageLoginPanelProperty"));
					WebElement loginPanel=driver.findElement(loginPanelProperty);
						String loginPanelText=loginPanel.getText();
				
				Cell Actual_HomePageText=row.createCell(10);
				Actual_HomePageText.setCellValue(loginPanelText);
				
				String expected_HomePageText=row.getCell(9).getStringCellValue();
				Log.info("The Expected Home Page Text is :- "+expected_HomePageText);
				
				String actual_HomePageText=row.getCell(10).getStringCellValue();
				Log.info("The Actual Home Page Text is:- "+actual_HomePageText);
				if(actual_HomePageText.equals(expected_HomePageText))
				{
					Log.info("The Actual Home Page Text is same as Expected Home Page Text - Pass");
						Cell TextTestResult=row.createCell(11);
						TextTestResult.setCellValue("Pass");
				}
				else
				{
					Log.info("The Actual Home Page Text is Not same as Expected Home Page Text - Fail");
						Cell TextTestResult=row.createCell(11);
						TextTestResult.setCellValue("Fail");
				}
				
				By spanMessageProperty=By.id(properties.getProperty("orangeHRMApplicationLogInPageSpanMessageProperty"));
					WebElement spanMessage=driver.findElement(spanMessageProperty);
						String spanMesssageText=spanMessage.getText();
				
				Cell Actual_LogInPageText=row.createCell(2);
				Actual_LogInPageText.setCellValue(spanMesssageText);
				
				String expected_loginPageTextOfCell1=row.getCell(1).getStringCellValue();
				Log.info("The Expected span Message Text is:- "+expected_loginPageTextOfCell1);
				
				String actual_LoginPageText=row.getCell(2).getStringCellValue();
				Log.info("The Actual Span Message Text is:- "+actual_LoginPageText);
				
				if(actual_LoginPageText.equals(expected_loginPageTextOfCell1))
				{
					Log.info("The Actual Span Message is same as Expected Span Message - Pass");
						Cell TextTestResult=row.createCell(3);
						TextTestResult.setCellValue("Pass");
				}
				else
				{
					Log.info("The Actual Span Message is Not same as Expected Span Message - Fail");
						Cell TextTestResult=row.createCell(3);
						TextTestResult.setCellValue("Fail");
				}
				
			}
		
		}
		FileOutputStream testDataResult=new FileOutputStream("./src/main/java/com/OrangeHRMApplicationTestResultFiles/OrangeHRMApplicationLoginTestDataResult.xlsx");
		workBook.write(testDataResult);
	}


}
