package DriverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunctionLibrary.FunctionLibrary;
import Utilities.ExcelFileUtil;

public class DriverScript
{
	WebDriver driver;
	ExtentReports report;
	ExtentTest logger;
	
	public void startTest() throws Throwable
	{
		ExcelFileUtil excel=new ExcelFileUtil();
		//MasterTestCase Sheet
		for (int i = 1; i <=excel.rowCount("MasterTestCase"); i++)
		{
			String ModuleStatus="";
					
			if(excel.getData("MasterTestCase",i, 2).equalsIgnoreCase("Y"))
			{
				//Define Module Name
				String TCModule=excel.getData("MasterTestCase", i, 1);
				report=new ExtentReports("D:\\ESTHAN\\AUTOMATION\\Automation Testing\\MvnStock\\Reports\\"+TCModule+FunctionLibrary.generateDate()+".html");
				logger=report.startTest(TCModule);
				
				int rowcount=excel.rowCount(TCModule);
				//TCModule sheet
				for (int j = 1; j <=rowcount; j++)
				{
					String Description=excel.getData(TCModule, j, 0);
					String Object_Type=excel.getData(TCModule, j, 1);
					String Locator_Type=excel.getData(TCModule, j, 2);
					String Locator_Value=excel.getData(TCModule, j, 3);
					String Test_Data=excel.getData(TCModule, j, 4);
					//System.out.println(Object_Type);
					
					try
					{
					if(Object_Type.equalsIgnoreCase("startBrowser"))
					{
						driver=FunctionLibrary.startBrowser(driver);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("openApplication"))
					{
						FunctionLibrary.openApplication(driver);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("typeAction"))
					{
						FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("clickAction"))
					{
						FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("closeBrowser"))
					{
						FunctionLibrary.closeBrowser(driver);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("waitForElement"))
					{
						FunctionLibrary.waitForElement(driver, Locator_Type, Locator_Value, Test_Data);
						logger.log(LogStatus.INFO, Description);
					} 
					if(Object_Type.equalsIgnoreCase("pageDown"))
					{
						FunctionLibrary.pageDown(driver);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("captureData"))
					{
						FunctionLibrary.captureData(driver, Locator_Type, Locator_Value);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("tableValidation"))
					{
						FunctionLibrary.tableValidation(driver, Test_Data);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("mouseClick"))
					{
						FunctionLibrary.mouseClick(driver, Locator_Type, Locator_Value);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("tableValidation1"))
					{
						FunctionLibrary.tableValidation1(driver, Test_Data);
						logger.log(LogStatus.INFO, Description);
					}
					//Status update TCModuleSheet pass
					excel.setData(TCModule, j, 5, "Pass");
					logger.log(LogStatus.PASS, Description+"    PASS");
					ModuleStatus="true";
					
					}
					catch(Exception e)
					{
						excel.setData(TCModule, j, 5, "Fail");
						logger.log(LogStatus.FAIL, Description+"  FAIL");
						ModuleStatus="false";
			File srcFile= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(srcFile, new File
		("D:\\ESTHAN\\AUTOMATION\\Automation Testing\\MvnStock\\ScreenShot\\"+TCModule+" "+
			FunctionLibrary.generateDate()+".png"));
						break;
						
						
					}
				}   
				if(ModuleStatus.equalsIgnoreCase("true"))
				{
					//status update MasterTestCase Pass
					excel.setData("MasterTestCase", i, 3, "Pass");
				}else
					if(ModuleStatus.equalsIgnoreCase("false"))
					{
						//status update in MasterTestCase Sheet "fail"
						excel.setData("MasterTestCase", i, 3, "Fail");
					}
				report.endTest(logger);
				report.flush();
				
				
			}else
			{
				excel.setData("MasterTestCase", i, 3, "Not Executed");
			}
			
		}
	}

}
