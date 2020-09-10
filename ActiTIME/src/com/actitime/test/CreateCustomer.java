package com.actitime.test;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.testng.Assert;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.actitime.genirics.BaseClass;
import com.actitime.pom.HomePage;
import com.actitime.pom.TaskListPage;

//@Listeners(com.actitime.genirics.ListenerImplementation.class)
public class CreateCustomer extends BaseClass {
	@Test
	public void testCreateCustomer() throws InterruptedException, EncryptedDocumentException, IOException {
		String customerName = fl.getExcelValue("CreateCustomer", 1, 0);
		String description = fl.getExcelValue("CreateCustomer", 1, 1);
		
		HomePage hm = new HomePage(driver);
		hm.clickTask();

		TaskListPage t = new TaskListPage(driver);
		t.getAddNewButton().click();
		t.getAddmenuCustomer().click();
		t.getEnterCustomerName().sendKeys(customerName);
		t.getEnterCustomerDescription().sendKeys(description);
		t.getDropDownOnProject().click();
		t.getOurCompany().click();
		t.getCreateCustomerButton().click();
		Thread.sleep(10000);
		String actualText = t.getActualCustomerCreated().getText();
		Assert.assertEquals(actualText, customerName);
	}
}

