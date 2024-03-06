package com.OrangeHRMApplicationTestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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
import com.Utility.Log;

public class OrangeHRMApplication_AddEmployeeTest extends BaseTestOrangeHRMApplication

{
	FileInputStream testDataFile;
	XSSFWorkbook workBook;
	XSSFSheet testDataSheet;
	
	@Test(priority = 1,description = "Adding Test Data into Excel Sheet")
	public void ExcelTestData() throws IOException
	{
		testDataFile=new FileInputStream("./src/main/java/com/OrangeHRMApplicationTestDataFiles/AddEmployeeExcelSheet.xlsx");
		workBook=new XSSFWorkbook(testDataFile);
		testDataSheet=workBook.getSheet("Sheet1");
		
		Row row1=testDataSheet.createRow(1);
		
		Cell Expected_loginPageTextOfCell0=row1.createCell(0);
		Expected_loginPageTextOfCell0.setCellValue("LOGIN Panel");
		
		Cell Expected_LoginPageTextOfCell1=row1.createCell(1);
		Expected_LoginPageTextOfCell1.setCellValue("Invalid credentials");
		
		Cell Expected_LogInPageTitle=row1.createCell(4);
		Expected_LogInPageTitle.setCellValue("orangeHRM");
		
		
		Cell UserName=row1.createCell(7);
		UserName.setCellValue("Admin");
		
		Cell Password=row1.createCell(8);
		Password.setCellValue("Chaitu@022");
		
		Cell Expected_HomePageText=row1.createCell(9);
		Expected_HomePageText.setCellValue("Admin");
		
		Cell Expected_AddEmployeePageText=row1.createCell(12);
		Expected_AddEmployeePageText.setCellValue("Add Employee");
		
		Cell Expected_FirstName=row1.createCell(15);
		Expected_FirstName.setCellValue("Vayu");
		
		Cell Expected_MiddleName=row1.createCell(16);
		Expected_MiddleName.setCellValue("Putra");
		
		Cell Expected_LastName=row1.createCell(17);
		Expected_LastName.setCellValue("Hanuma");
		
		
		Cell Expected_LogInPageTextOfCell30=row1.createCell(30);
		Expected_LogInPageTextOfCell30.setCellValue("LOGIN Panel");		
		// creating second Row
		
		Row row2=testDataSheet.createRow(2);
		
		Expected_FirstName=row2.createCell(15);
		Expected_FirstName.setCellValue("Kesari");
		
		Expected_MiddleName=row2.createCell(16);
		Expected_MiddleName.setCellValue("Nandana");
		
		Expected_LastName=row2.createCell(17);
		Expected_LastName.setCellValue("Hanuma");
		
		// creating third Row
		
		Row row3=testDataSheet.createRow(3);
		
		Expected_FirstName=row3.createCell(15);
		Expected_FirstName.setCellValue("Rama");
		
		Expected_MiddleName=row3.createCell(16);
		Expected_MiddleName.setCellValue("Dhuta");
		
		Expected_LastName=row3.createCell(17);
		Expected_LastName.setCellValue("Hanuma");

	
	}
	

	@Test(priority = 2,description = "Validating OrangeHRM Application LogIn Page Title")
	public void Validating_LoginPageTitle()
	{
		Row testDataRow=testDataSheet.getRow(1);
		
		String expected_LogInPageTitle=testDataRow.getCell(4).getStringCellValue();
		Log.info("The Expected Login Page Title is :-"+expected_LogInPageTitle);
		
		String actual_LogInPageTitle=driver.getTitle();
		Log.info("The Actual Login Page Title is :- "+actual_LogInPageTitle);
		
		Cell Actual_LogInPageTitle=testDataRow.createCell(5);
		Actual_LogInPageTitle.setCellValue(actual_LogInPageTitle);
		
		if(actual_LogInPageTitle.equals(expected_LogInPageTitle))
		{
			Log.info("The Actual LoginPage Title is same as Expected Loginpage Title - Pass");
			
			Cell TitleTest_Result=testDataRow.createCell(6);
			TitleTest_Result.setCellValue("Pass");
		}
		else
		{
			Log.info("The Actual LoginPage Title is Not same as Expected Loginpage Title - Fail");
			
			Cell TitleTest_Result=testDataRow.createCell(6);
			TitleTest_Result.setCellValue("Fail");
		}	
	}
	
	
	
	FileInputStream propertiesFile;
	Properties properties;
	
	@Test(priority = 3,description = "validating the OrangeHRM Aplication Login Page LoGo")
	public void ValidatingLogo() throws IOException 
	{
		propertiesFile=new FileInputStream("./src/main/java/com/Config/OrangeHRMApplication_AddEmployeeTest.properties");
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
	
	@Test(priority = 4,description = "Validating OrangeHRM Application LogIn Functionality Test ")
	public void Validating_loginTest() throws IOException
	{
		 
		
		Row testDataRow=testDataSheet.getRow(1);
		
		String userNameTextData=testDataRow.getCell(7).getStringCellValue();
			By usernameProperty=By.id(properties.getProperty("orangeHRMApplicationLogInPageUserNameProperty"));
				WebElement userName = driver.findElement(usernameProperty);
					userName.sendKeys(userNameTextData);
		
		String passwordTextData=testDataRow.getCell(8).getStringCellValue();
			By passwordProperty=By.id(properties.getProperty("orangeHRMApplicatonLogInPagePasswordProperty"));
				WebElement password = driver.findElement(passwordProperty);
					password.sendKeys(passwordTextData);
		
		By loginButtonProperty=By.id(properties.getProperty("orangeHRMApplicationLogInPagebuttonProperty"));
			WebElement loginButton = driver.findElement(loginButtonProperty);
				loginButton.click();	
				
	}
	
	@Test(priority = 5,description = "Identifying the Home Page PiM")
	public void Validating_PimTest() 
	{
		
		By pimProperty=By.id(properties.getProperty("orangeHRMApplicationHomePagePimProperty"));
		WebElement pim = driver.findElement(pimProperty);
		
		
		Actions ActionName=new Actions(driver);
		ActionName.moveToElement(pim).build().perform();
	}
	
	@Test(priority = 6,description = "Identifying the Home Page PIM Add Employee")
	public void Validating_AddEmployeeTest()
	{
		By AddEmployeeProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageAddEmployeeProperty"));
		WebElement AddEmployee = driver.findElement(AddEmployeeProperty);
			AddEmployee.click();
	}
	
	@Test(priority = 7,description = "Validating OrangeHRM Application Home Page pim Add Employee Text")
	public void Validating_AddEmployeeText() 	{
		Row row=testDataSheet.getRow(1);
		
		By addEmployeePageTxtProperty=By.xpath(properties.getProperty("orangeHRMApplicationHomePageAddEmployeePageTextProperty"));
			WebElement addEmployeePageTxt = driver.findElement(addEmployeePageTxtProperty);
				String addEmployeePageText=addEmployeePageTxt.getText();
		

		Cell Actual_AddEmployeePageText =row.createCell(13);
		Actual_AddEmployeePageText.setCellValue(addEmployeePageText);	
			
		String expected_AddEmployeePageText=row.getCell(12).getStringCellValue();
		Log.info("The Expected Add Employee Page Text is:- "+expected_AddEmployeePageText);
		
		String actual_AddEmployeePageText=row.getCell(13).getStringCellValue();
		Log.info("The Actual Add Employee page Text is:- "+actual_AddEmployeePageText);
		if(actual_AddEmployeePageText.equals(expected_AddEmployeePageText))
		{
			Log.info("Text for the Add Employee Page is Matched -Pass");
			Cell AddEmployeeTextTest_Result=row.getCell(14);
			AddEmployeeTextTest_Result.setCellValue("Pass");

		}
		else
		{
			Log.info("Text for the Add Employee Page is Not Matched -Fail");
			Cell AddEmployeeTextTest_Result=row.getCell(14);
			AddEmployeeTextTest_Result.setCellValue("Fail");
		}
	
	}
	
	@Test(priority = 8,description = "Filling Excel Sheet with Actual Add Employee Page Details")
	public void Adding_AddEmployeeTestData() throws IOException, InterruptedException
	{

		int rowCount=testDataSheet.getLastRowNum();
		
		for(int rowIndex=1;rowIndex<=rowCount;rowIndex++)
		{
			Row row=testDataSheet.getRow(rowIndex);
			
			
			String firstNameTestData=row.getCell(15).getStringCellValue();
				By firstNameProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageAddEmployeePageFirstNameProperty"));
					WebElement firstName=driver.findElement(firstNameProperty);
						firstName.sendKeys(firstNameTestData);			
				
			String middleNameTestData=row.getCell(16).getStringCellValue(); 
			
			Actions keyboardActions = new Actions(driver);
			keyboardActions.sendKeys(Keys.TAB).build().perform();
			
			keyboardActions.sendKeys(middleNameTestData).build().perform();
			
			
			String lastNameTestData=row.getCell(17).getStringCellValue();
		
			keyboardActions.sendKeys(Keys.TAB).build().perform();
			keyboardActions.sendKeys(lastNameTestData).build().perform();
		
								
			By employeeIdProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageAddEmployeePageEmployeeIdProperty"));
				WebElement employeeid = driver.findElement(employeeIdProperty);
					String employeeId_Text=employeeid.getAttribute("value");
			
			Log.info("The Add Employee EmployeeId Value is:- "+ employeeId_Text);
					
			Cell Expected_Employee_ID=row.createCell(18);
			Expected_Employee_ID.setCellValue(employeeId_Text);
			
			keyboardActions.sendKeys(Keys.TAB).build().perform();
			keyboardActions.sendKeys(Keys.TAB).build().perform();
			
			;
			keyboardActions.sendKeys(Keys.ENTER).build().perform();
			
			Thread.sleep(9000);
		
			
			java.lang.Runtime.getRuntime().exec("./AutoITTestScripts/OrangeHRMAddEmployeeCompileScriptFile.exe");
			
			Thread.sleep(5000);
					
			By saveButtonProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageAddEmployeePageSaveButtonProperty"));
				WebElement saveButton=driver.findElement(saveButtonProperty);
					saveButton.click();
					
			
					
			Cell Expected_PersonalDetailsPageText=row.createCell(19);
			Expected_PersonalDetailsPageText.setCellValue("Personal Details");
			
			By PersonalDetailsProperty =By.xpath(properties.getProperty("orangeHRMApplicationHomePagePersonalDetailsPageTextProperty"));
				WebElement PersonalDetailsTxt=driver.findElement(PersonalDetailsProperty);
					String PersonalDetailsPageText=PersonalDetailsTxt.getText();
			
			Cell Actual_PersonalDetailsPageText=row.createCell(20);
			Actual_PersonalDetailsPageText.setCellValue(PersonalDetailsPageText);
			
			By employeeListFirstNameProperty=By.name(properties.getProperty("orangeHRMApplicationHomePageEmployeeListPersonalDetailsPageFirstNameProperty"));
				WebElement employeeListFirstName = driver.findElement(employeeListFirstNameProperty);
					String 	firstNameValue=employeeListFirstName.getAttribute("value");
						
			Cell Actual_FirstName=row.createCell(22);
			Actual_FirstName.setCellValue(firstNameValue);
			
			By employeeListMiddleNamePrperty=By.name(properties.getProperty("orangeHRMApplicationHomePageEmployeeListPersonalDetailsPageMiddleNameProperty"));
				WebElement employeeListMiddleName=driver.findElement(employeeListMiddleNamePrperty);
					String middleNameValue=employeeListMiddleName.getAttribute("value");
			
			Cell Actual_MiddleName=row.createCell(24);
			Actual_MiddleName.setCellValue(middleNameValue);
			
			By employeeListLastNamePrperty=By.name(properties.getProperty("orangeHRMApplicationHomePageEmployeeListPersonalDetailsPageLastNameProperty"));
				WebElement employeeListLastName=driver.findElement(employeeListLastNamePrperty);
					String lastNameValue=employeeListLastName.getAttribute("value");
		
			Cell Actual_LastName=row.createCell(26);
			Actual_LastName.setCellValue(lastNameValue);
			
			By employeeListEmpIdPrperty=By.name(properties.getProperty("orangeHRMApplicationHomePageEmployeeListPersonalDetailsPageEmployeeIdProperty"));
				WebElement employeeListEmpId=driver.findElement(employeeListEmpIdPrperty);
					String employeeListEmpIdValue=employeeListEmpId.getAttribute("value");
	
			Cell Actual_EmployeeId=row.createCell(28);
			Actual_EmployeeId.setCellValue(employeeListEmpIdValue);			
		
			By AddEmployeeProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageAddEmployeeProperty"));
				WebElement AddEmployee = driver.findElement(AddEmployeeProperty);
					AddEmployee.click();
					
		}	
		

	}
	

    @Test(priority = 9,description = "Validating the HomePage Text and LoginPage Text")
	
	public void validatingHomePageAndLoginPage_Test()
	{
		
		
		Row row=testDataSheet.getRow(1);
		
		By AdminProperty=By.id(properties.getProperty("orangeHRMApplicationHomePageAdminProperty"));
			WebElement Admin=driver.findElement(AdminProperty);
				String AdminText=Admin.getText();  
		
		Cell Actual_HomePageText=row.createCell(10);
		Actual_HomePageText.setCellValue(AdminText);
			
		String expected_HomePageText=row.getCell(9).getStringCellValue();
		Log.info("The Expected HomePage Text is :- "+expected_HomePageText);
		
		String actual_HomePageText=row.getCell(10).getStringCellValue();
		Log.info("The Actual HomePage Text is :- "+actual_HomePageText);

		
		if(actual_HomePageText.equals(expected_HomePageText))
		{
			Log.info("The Actual HomePage Text is same as The Expected HomePage Text is Pass");
			
			Cell HomePageTest_Result=row.createCell(11);
			HomePageTest_Result.setCellValue("Pass");
				
				Cell Actual_LogInPageText=row.createCell(2);
				Actual_LogInPageText.setCellValue("Valid Credentials");
	    		
				String expected_loginText=row.getCell(1).getStringCellValue();
				Log.info("The Expected loginPage is :-"+expected_loginText);
	    			
				String actual_loginText=row.getCell(2).getStringCellValue();
				Log.info("The Actual loginPage is :-"+actual_loginText);

	    			if(actual_loginText.equals(expected_loginText))
	    			{
	    				Log.info("The Credentials are pass");
	    				Cell RowOfCell03=row.createCell(3);
	    				RowOfCell03.setCellValue("Pass");
	    			}
	    			else
	    			{
	    				Log.info("The Credentials are Fail");
	    		 
	    				Cell RowOfCell11=row.createCell(3);
	    				RowOfCell11.setCellValue("Fail");
	    			}
			}
			else
			{
				Log.info("The actual HomePage Text is Not same as the Expected HomePage Text is Fail ");
					Cell HomePageTest_Result=row.createCell(11);
					HomePageTest_Result.setCellValue("Fail");
			}
			
	}
    
   @Test(priority = 10,description = "Filling Excel Sheet With Add Employee Page Results")
	public void Validating_AddEmployeePage() throws IOException
	{
		
		int rowCount=testDataSheet.getLastRowNum();
		for(int rowindex=1;rowindex<=rowCount;rowindex++)
		{
			Row row=testDataSheet.getRow(rowindex);
			
				
				// Validating the OrangeHRM Application Personal Details Page Text
					String expected_PersonalDetailsPageText=row.getCell(19).getStringCellValue();
					Log.info("The Expected Personal Details Page Text is :- "+expected_PersonalDetailsPageText);
					
					String actual_PersonalDetailsPageText=row.getCell(20).getStringCellValue();
					Log.info("The Actual Personal Details Page Text is :- "+actual_PersonalDetailsPageText);
					
					if(actual_PersonalDetailsPageText.equals(expected_PersonalDetailsPageText))
					{
						Log.info("The Actual Personal Details Page Text is same as Expected Personal Details Page Text - Pass");
							Cell PersonalDetailsPage_TestResult=row.createCell(21);
							PersonalDetailsPage_TestResult.setCellValue("Pass");
					}
					else
					{
						Log.info("The Actual Personal Details Page Text is Not same as Expected Personal Details Page Text - Fail");
							Cell PersonalDetailsPage_TestResult=row.createCell(21);
							PersonalDetailsPage_TestResult.setCellValue("Fail");
					}
		
			//Validating the OrangeHRM Application First Name 
				
			String expected_FirstName=row.getCell(15).getStringCellValue();
			Log.info("The Expected First Name is :- "+expected_FirstName);
				
			String actual_FirstName=row.getCell(22).getStringCellValue();
			Log.info("The Actual First Name is :- "+actual_FirstName);
				
			if(actual_FirstName.equals(expected_FirstName))
			{
				Log.info("The Actual First Name is same as the Expected First Name - Pass");
					Cell FirstName_TestResult=row.createCell(23);
					FirstName_TestResult.setCellValue("Pass");
			}
			else
			{
				Log.info("The Actual FIrst Name is Not same as the Expected First Name - Fail");
					Cell FirstName_TestResult=row.createCell(23);
					FirstName_TestResult.setCellValue("Fail");
			}
				
				
				//Validating the OrangeHRM Application Middle Name
					String expected_MiddleName=row.getCell(16).getStringCellValue();
					Log.info("The Expected Middle Name is :- "+expected_MiddleName);
						
					String actual_MiddleName=row.getCell(24).getStringCellValue();
					Log.info("The Actual Middle Name is :- "+actual_MiddleName);
						
					if(actual_MiddleName.equals(expected_MiddleName))
					{
						Log.info("The Actual Middle Name is same as the Expected Middle Name - Pass");
							Cell MiddleName_TestResult=row.createCell(25);
							MiddleName_TestResult.setCellValue("Pass");
					}
					else
					{
						Log.info("The Actual Middle Name is Not same as the Expected Middle Name - Fail");
							Cell MiddleName_TestResult=row.createCell(25);
							MiddleName_TestResult.setCellValue("Fail");
					}
						
			// Validating the OrangeHRM Application LastName 
			String expected_LastName=row.getCell(17).getStringCellValue();
			Log.info("The Expected First Name is :- "+expected_LastName);
				
			String actual_LastName=row.getCell(26).getStringCellValue();
			Log.info("The Actual First Name is :- "+actual_LastName);
				
			if(actual_LastName.equals(expected_LastName))
			{
				Log.info("The Actual Last Name is same as the Expected Last Name - Pass");
					Cell LastName_TestResult=row.createCell(27);
					LastName_TestResult.setCellValue("Pass");
			}
			else
			{
				Log.info("The Actual Last Name is Not same as the Expected Last Name - Fail");
					Cell LastName_TestResult=row.createCell(27);
					LastName_TestResult.setCellValue("Fail");
			}
				
				//Validating OrangeHRM Application Employee ID
					String expected_Employee_ID=row.getCell(18).getStringCellValue();
					Log.info("The Expected Employee ID is :- "+expected_Employee_ID);
		
					String actual_Employee_ID=row.getCell(28).getStringCellValue();
					Log.info("The Actual Employee ID is :- "+actual_Employee_ID);
						
					if(actual_Employee_ID.equals(expected_Employee_ID))
					{
						Log.info("The Actual Employee ID is same as the Expected Employee ID - Pass");
							Cell EmployeeID_Result=row.createCell(29);
							EmployeeID_Result.setCellValue("Pass");
					}
					else
					{
						Log.info("The Actual Employee ID is Not same as the Expected Employee ID - Fail");
							Cell EmployeeID_Result=row.createCell(29);
							EmployeeID_Result.setCellValue("Fail");
					}
						
			
			
		}
	}
    
    @Test(priority = 11,description = "OrangeHRM Application Home Page Logout Functionality")
	public void Validating_logOutTest() 
	{
		By welcomeAdminproperty=By.id(properties.getProperty("orangeHRMApplicationHomePageWelcomeAdminProperty"));
			WebElement welcomeAdmin=driver.findElement(welcomeAdminproperty);
				welcomeAdmin.click();
			
				
		By logOutProperty=By.linkText(properties.getProperty("orangeHRMApplicationHomePageLogOutProperty"));	
			WebElement logOut=driver.findElement(logOutProperty);
				logOut.click();
	}
    
    
    @Test(priority = 12,description = "Validating OrangeHRM Application LogOut Functionality Test")
	public void validating_LoginPage() throws IOException
	{
		Row row=testDataSheet.getRow(1);
		
		By loginPanelProperty=By.id(properties.getProperty("orangeHRMApplicationLoginPageLoginPanelProperty"));
			WebElement loginPanel=driver.findElement(loginPanelProperty);
				String logInPanelText=loginPanel.getText();
		
		Cell Actual_LoginPageText=row.createCell(31);		
		Actual_LoginPageText.setCellValue(logInPanelText);
		
		String expected_LoginPageText=row.getCell(30).getStringCellValue();
		Log.info("The Expected Login Page Text is :- "+expected_LoginPageText);
			
		String actual_LoginPageText=row.getCell(31).getStringCellValue();
		Log.info("The Actual Login Page Text is :- "+actual_LoginPageText);
		
		if(actual_LoginPageText.equals(expected_LoginPageText))
		{
			Log.info("The Actual Login Page Text is same as Expected Login Page Text - Pass");
			Cell Expected_LoginPageText=row.createCell(32);		
			Expected_LoginPageText.setCellValue("Pass");
		}
		else
		{
			Log.info("The Actual Login Page Text is Not same as Expected Login Page Text - Fail");
			Cell Expected_LoginPageText=row.createCell(32);		
			Expected_LoginPageText.setCellValue("Fail");
			
		}
		
		
		FileOutputStream testDataResult=new FileOutputStream("./src/main/java/com/OrangeHRMApplicationTestResultFiles/OrangeHRMApplication_AddEmployeeTestResult.xlsx");	
		workBook.write(testDataResult);
		
	}

	

}
