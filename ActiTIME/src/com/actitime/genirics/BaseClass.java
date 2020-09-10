package com.actitime.genirics;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;

import com.actitime.pom.HomePage;
import com.actitime.pom.LoginPage;

public class BaseClass {

	static {
		System.setProperty("webdriver.chrome.driver", "./driver/chromedriver.exe");
	}
	public static WebDriver driver;
	public FileLib fl = new FileLib();

	@BeforeClass
	public void openBrowser() throws IOException {
		driver = new ChromeDriver();
		String url = fl.getPropertyValue("url");
		// driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get(url);

	}

	@BeforeMethod
	public void login() throws IOException {

		String un = fl.getPropertyValue("un");
		String pwd = fl.getPropertyValue("pw");
		LoginPage l = new LoginPage(driver);
		l.getUntbx().sendKeys(un);
		l.getPatbx().sendKeys(pwd);
		l.getLgbtn().click();
	}

	@AfterMethod
	public void logout() throws InterruptedException {
		HomePage h = new HomePage(driver);
		Thread.sleep(10000);
		h.LogoutBtn();

	}

	@AfterClass
	public void closeBrowser() {
		driver.close();
	}

}
