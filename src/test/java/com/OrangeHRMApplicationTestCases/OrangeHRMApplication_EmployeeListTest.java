package com.OrangeHRMApplicationTestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import com.BaseTests.BaseTestOrangeHRMApplication;

public class OrangeHRMApplication_EmployeeListTest extends BaseTestOrangeHRMApplication
{
	FileInputStream propertiesFile;
	Properties properties;
	@Test(priority = 1,description = "OrangeHRM Application Login Page Login Funcationality")
	public void Validting_LoginTest() throws IOException
	{
		propertiesFile=new FileInputStream("./src/main/java/com/Config/OrangeHRMApplication_EmployeeListTest.properties");
		properties= new Properties();
		properties.load(propertiesFile);
		
		String userNameTestData="Admin";
			By userNameProperty=By.id(properties.getProperty("orangeHRMApplicationLogInPageUserNameProperty"));
			WebElement userName=driver.findElement(userNameProperty);
				userName.sendKeys(userNameTestData);
		
		String passwordTestData="Chaitu@022";
		Actions keyboardAction=new Actions(driver);
		keyboardAction.sendKeys(Keys.TAB).build().perform();
		keyboardAction.sendKeys(passwordTestData).build().perform();
		
		keyboardAction.sendKeys(Keys.TAB).build().perform();
		keyboardAction.sendKeys(Keys.ENTER).build().perform();
		
	}
	
	@Test(priority = 2,description = "Identifying the OrangeHRM Application Home Page MouseHovering to PIM ")
	public void Validating_PimTest()
	{
		By pimProperty=By.id(properties.getProperty("orangeHRMApplicationHomePagePimProperty"));
		WebElement pim=driver.findElement(pimProperty);
		
		Actions mouseAction=new Actions(driver);
		mouseAction.moveToElement(pim).build().perform();
		
	}
	
	@Test(priority = 3,description = "Identifing the OrangeHRM Application Home Page Click on EmployeeList ")
	public void Validating_EmployeeListTest()
	{
		By employeelistProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageEmployeeListProperty"));
		WebElement employeelist=driver.findElement(employeelistProperty);
		employeelist.click();
	}
	
	@Test(priority = 4,description = "Identifying the Employee List Details")
	public void Validating_EmployeeListDetails() throws IOException
	{
		FileInputStream testDataFile=new FileInputStream("./src/main/java/com/OrangeHRMApplicationTestDataFiles/EmployeeListEmptySheet.xlsx");
		XSSFWorkbook workBook=new XSSFWorkbook(testDataFile);
		XSSFSheet testDataSheet=workBook.getSheet("TestData");
		
		///html/body/div[1]/div[3]/div[2]/div/form/div[1]/ul/li[6]/a
		By tablePageProperty=By.xpath("/html/body/div[1]/div[3]/div[2]/div/form/div[1]/ul/li[6]/a");
		WebElement tablePage=driver.findElement(tablePageProperty);
		tablePage.click();
		
		By tableBodyProperty=By.xpath(properties.getProperty("orangeHRMApplicationEmployeeListTableBodyProperty"));
		WebElement tableBody=driver.findElement(tableBodyProperty);
		
		By tableRowProperty=By.tagName(properties.getProperty("orangeHRMApplicationEmployeeListTableRowsProperty"));
		List<WebElement> tableRow=tableBody.findElements(tableRowProperty);
	
		for(int rowIndex=0;rowIndex<tableRow.size();rowIndex++)
		{
			WebElement webTableRow=tableRow.get(rowIndex);
			Row row=testDataSheet.createRow(rowIndex);
			
			By tableCellProperty=By.tagName(properties.getProperty("orangeHRMApplicationEmployeeListTableCellsProperty"));
			List<WebElement> tableCell = webTableRow.findElements(tableCellProperty);
			
			for(int cellIndex=0;cellIndex<tableCell.size();cellIndex++)
			{
				WebElement webTableCell=tableCell.get(cellIndex);
				
				String webTableData = webTableCell.getText();
				
			
				Cell cell=row.createCell(cellIndex);
				cell.setCellValue(webTableData);
			}
		
		}
		FileOutputStream testDataResult=new FileOutputStream("./src/main/java/com/OrangeHRMApplicationTestResultFiles/OrangeHRMApplication_EmployeeListTestResult.xlsx");
		workBook.write(testDataResult);
	}

}
