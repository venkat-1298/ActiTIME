package com.actitime.pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.server.handler.SendKeys;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class LoginPage {
	@FindBy(id = "username")
	private WebElement untbx;
	@FindBy(name = "pwd")
	private WebElement patbx;
	@FindBy(xpath = "//div[text()='Login ']")
	private WebElement lgbtn;

	public LoginPage(WebDriver driver) {
		PageFactory.initElements(driver, this);
	}

	public WebElement getUntbx() {
		return untbx;
	}

	public WebElement getPatbx() {
		return patbx;
	}

	public WebElement getLgbtn() {
		return lgbtn;
	}

}
