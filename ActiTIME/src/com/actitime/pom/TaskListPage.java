package com.actitime.pom;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class TaskListPage {

	public TaskListPage(WebDriver driver) {
		PageFactory.initElements(driver, this);
	}

	@FindBy(xpath = "//div[text()='Add New']")
	private WebElement addNewButton;

	@FindBy(xpath = "//div[text()='+ New Customer']")
	private WebElement addmenuCustomer;

	@FindBy(xpath = "(//input[@placeholder='Enter Customer Name'])[2]")
	private WebElement enterCustomerName;

	@FindBy(xpath = "//textarea[@placeholder='Enter Customer Description']")
	private WebElement enterCustomerDescription;

	@FindBy(xpath = "(//div[@class='dropdownButton']) [15]")
	private WebElement dropDownOnProject;

	@FindBy(xpath = "//div[@class='itemRow cpItemRow ' and text()='Our company']")
	private WebElement ourCompany;

	@FindBy(xpath = "//div[text()='Create Customer']")
	private WebElement createCustomerButton;

	@FindBy(xpath = "(//div[@class='title' and @title]) [1]")
	private WebElement actualCustomerCreated;

	public WebElement getAddNewButton() {
		return addNewButton;
	}

	public WebElement getAddmenuCustomer() {
		return addmenuCustomer;
	}

	public WebElement getEnterCustomerName() {
		return enterCustomerName;
	}

	public WebElement getEnterCustomerDescription() {
		return enterCustomerDescription;
	}

	public WebElement getDropDownOnProject() {
		return dropDownOnProject;
	}

	public WebElement getOurCompany() {
		return ourCompany;
	}

	public WebElement getCreateCustomerButton() {
		return createCustomerButton;
	}

	public WebElement getActualCustomerCreated() {
		return actualCustomerCreated;
	}

}
