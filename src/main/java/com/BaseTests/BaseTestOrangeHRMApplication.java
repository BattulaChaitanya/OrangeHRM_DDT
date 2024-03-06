package com.BaseTests;


import java.io.IOException;
import java.util.concurrent.TimeUnit;


import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

import com.Utility.Log;


public class BaseTestOrangeHRMApplication 
{
	
	public WebDriver driver;
	String  applicationUrlAddress="http://127.0.0.1/orangehrm-4.2.0.1/symfony/web/index.php/auth/login";

	
	@BeforeTest
	public void setUp() throws IOException
	{
		
		System.setProperty("webdriver.chrome.driver","./BrowserFiles/chromedriver.exe");
		
		driver=new ChromeDriver();
		Log.info("********Chrome Browser Launched Successfully********");
		
		driver.navigate().to(applicationUrlAddress);
		
		Log.info("Navigated to OrangeHRM Application WebPage");
		driver.manage().window().maximize();
		
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
	}
	
	@AfterTest
	public void tearDown()
	{
		driver.quit();
		Log.info("********Chrome Browser Closed Successfully********");
	
		
	}

	

}
