package com.actitime.pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class HomePage {

	@FindBy(id = "logoutLink")
	private WebElement logout;

	@FindBy(id = "container_tasks")
	private WebElement taskTab;

	public HomePage(WebDriver driver) {
		PageFactory.initElements(driver, this);
	}

	public void LogoutBtn() throws InterruptedException {
		logout.click();
	}

	public void clickTask() {
		taskTab.click();
	}
}
