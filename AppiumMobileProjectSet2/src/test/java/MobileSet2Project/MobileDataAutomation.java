package MobileSet2Project;

import org.testng.annotations.Test;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.remote.MobileCapabilityType;

import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeClass;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.AfterClass;

public class MobileDataAutomation {
	public static WebDriver driver;
	 public static HSSFWorkbook workbook;
	 public static HSSFSheet worksheet;
	    public static DataFormatter formatter= new DataFormatter();
	    static String SheetName= "Sheet1";
	    public static String file_location = System.getProperty("user.dir")+"/TestData.xls";
  @Test(dataProvider = "ReadVariant")
  public void AddVariants(String url,String SortBy) throws InterruptedException {
	  driver.navigate().to(url);
	  Thread.sleep(4000);
	  driver.findElement(By.xpath("//span[@class='icon']")).click();
	  driver.findElement(By.xpath("//ul[@class='mob-top-menu show']//a[@href='/computers']")).click();
	  driver.findElement(By.xpath("//h2[@class='title']//a[@href='/desktops']")).click();
	  Select sort = new Select(driver.findElement(By.name("products-orderby")));
	  sort.selectByVisibleText(SortBy);
	  Select sortnew = new Select(driver.findElement(By.name("products-orderby")));
	  System.out.println(sortnew.getFirstSelectedOption().getText());
      Assert.assertEquals(sortnew.getFirstSelectedOption().getText(), SortBy, "This is to verify sorting");
  }

  @DataProvider
  public Object[][] ReadVariant() throws IOException {
	 
	FileInputStream fileInputStream= new FileInputStream(file_location); //Excel sheet file location get mentioned here
      HSSFWorkbook workbook = new HSSFWorkbook (fileInputStream); //get my workbook 
      worksheet=workbook.getSheet(SheetName);// get my sheet from workbook
      HSSFRow Row=worksheet.getRow(0);     //get my Row which start from 0   
   
      int RowNum = worksheet.getPhysicalNumberOfRows();// count my number of Rows
      int ColNum= Row.getLastCellNum(); // get last ColNum 
       
      Object Data[][]= new Object[RowNum-1][ColNum]; // pass my  count data in array
       
          for(int i=0; i<RowNum-1; i++) //Loop work for Rows
          {  
              HSSFRow row= worksheet.getRow(i+1);
               
              for (int j=0; j<ColNum; j++) //Loop work for colNum
              {
                  if(row==null)
                      Data[i][j]= "";
                  else
                  {
                      HSSFCell cell= row.getCell(j);
                      if(cell==null)
                          Data[i][j]= ""; //if it get Null value it pass no data 
                      else
                      {
                          String value=formatter.formatCellValue(cell);
                          Data[i][j]=value; //This formatter get my all values as string i.e integer, float all type data value
                      }
                  }
              }
          }

      return Data;
  }
  
  @BeforeClass
  public void beforeClass() throws MalformedURLException {
	  DesiredCapabilities capability= new DesiredCapabilities();     
	  capability.setCapability(MobileCapabilityType.DEVICE_NAME, "Nexus 5x");       
	  capability.setCapability(MobileCapabilityType.PLATFORM_NAME, "Android");       
	  capability.setCapability(MobileCapabilityType.BROWSER_NAME, "Chrome");        
	  driver=new AndroidDriver(new URL("http://0.0.0.0:4723/wd/hub"),capability);
  }

  @AfterClass
  public void afterClass() {
	  driver.close();
	  driver.quit();
  }

}
