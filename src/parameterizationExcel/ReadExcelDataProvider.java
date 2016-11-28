package parameterizationExcel;

import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.testng.annotations.DataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;

public class ReadExcelDataProvider {
	
	public WebDriver driver;
	public WebDriverWait myWaitVar; 
	String appURL = "http://carhire.com.au/";
	
	//Locators
	private By byPickupLoc = By.id("pickupLoc");
	private By bySearch = By.cssSelector("div.ui.blue.large.submit.button.ng-isolate-scope");
	private By byError = By.cssSelector("div.ui.error.message");
	private By resultCount= By.cssSelector(".resultsCount > strong");

	  @BeforeClass
	  public void beforeTest() {
		  
		  System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
			driver=new ChromeDriver();
			driver.manage().window().maximize();
			myWaitVar = new WebDriverWait(driver, 30);
	  }	
	
  @Test(dataProvider="searchLocation")
  public void VerifyValidLocations(String topLocation)throws InterruptedException {
	  
	
	    driver.navigate().to(appURL);
		myWaitVar.until(ExpectedConditions.visibilityOfElementLocated(byPickupLoc));
		driver.findElement(byPickupLoc).sendKeys(topLocation);
		myWaitVar.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("strong.ng-binding")));
		driver.findElement(By.cssSelector("strong.ng-binding")).click();
		
		//wait for element to be visible and perform click
		myWaitVar.until(ExpectedConditions.visibilityOfElementLocated(bySearch));
		driver.findElement(bySearch).click();
		
			
		//Results Page
		myWaitVar.until(ExpectedConditions.visibilityOfElementLocated(resultCount));
		
		String counter= driver.findElement(resultCount).getText();
		System.out.println("Result Count of "+topLocation+" = "+counter);
		
		/*if(driver.getPageSource().contains("ui small expanded-location modal bpadding10 cf transition visible active scrollings")){
			
			
			Assert.fail();
		}
		
		else{
			
			Assert.assertTrue(true);
		} */
			
		
		/*if (driver.getPageSource().contains("ui small expanded-location modal bpadding10 cf transition visible active scrollings")){		
	    Assert.fail();
		}*/
  }


  @AfterClass
  public void afterTest() {
	  
	  driver.quit();
  }
  
  @DataProvider(name="searchLocation")
	public Object[][] searchData() {
		Object[][] arrayObject = getExcelData("E:/ODESK/VroomVroomVroom/Top Locations/Top Locations 4401-4878.xls","Sheet1");
		return arrayObject;
	}
  
  /**
	 * @param File Name
	 * @param Sheet Name
	 * @return
	 */
  
  public String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);

			int totalNoOfCols = sh.getColumns();
			int totalNoOfRows = sh.getRows();
			
			arrayExcelData = new String[totalNoOfRows-1][totalNoOfCols];
			
			for (int i= 1 ; i < totalNoOfRows; i++) {

				for (int j=0; j < totalNoOfCols; j++) {
					arrayExcelData[i-1][j] = sh.getCell(j, i).getContents();
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}

}

