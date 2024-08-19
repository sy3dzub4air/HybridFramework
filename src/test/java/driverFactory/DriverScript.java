package driverFactory;

import org.openqa.selenium.WebDriver;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript {
	WebDriver driver;
	String inputpath ="./FileInput/DataEngine.xlsx";
	String outputpath ="./FileOutput/HybridResults.xlsx";
	ExtentReports report;
	ExtentTest logger;
	String TestCases ="MasterTestCases";
	public void startTest() throws Throwable
	{
		String Module_Status;
		//create object for Exelfileutil class
		ExcelFileUtil xl = new ExcelFileUtil(inputpath);
		//iterate all test cases in TestCases
		for(int i=1;i<=xl.rowCount(TestCases);i++)
		{
			if(xl.getCellData(TestCases, i, 2).equalsIgnoreCase("Y"))
			{
				//read corresponding sheet or test cases
				String TCModule =xl.getCellData(TestCases, i, 1);
				//define path of html 
				report= new ExtentReports("./target/ExtentReports/"+TCModule+FunctionLibrary.generateDate()+".html");
				logger= report.startTest(TCModule);
				logger.assignAuthor("Syed Zubair");

				//iterate all rows in TCModule sheet
				for(int j=1;j<=xl.rowCount(TCModule);j++)
				{
					//read all cell from TCmodule sheet
					String Description = xl.getCellData(TCModule, j, 0);
					String Object_Type = xl.getCellData(TCModule, j, 1);
					String LName =xl.getCellData(TCModule, j, 2);
					String Lvalue = xl.getCellData(TCModule, j, 3);
					String TestData = xl.getCellData(TCModule, j, 4);
					try {
						if(Object_Type.equalsIgnoreCase("startBrowser"))
						{
							driver =FunctionLibrary.startBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("openUrl"))
						{
							FunctionLibrary.openUrl();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("waitForElement"))
						{
							FunctionLibrary.waitForElement(LName, Lvalue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(LName, Lvalue, TestData);
							Thread.sleep(20);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(LName, Lvalue);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("validateTitle"))
						{
							FunctionLibrary.validateTitle(TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("mouseClick"))
						{
							FunctionLibrary.mouseClick();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("categoryTable"))
						{
							FunctionLibrary.categoryTable(TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("dropDownAction"))
						{
							FunctionLibrary.dropDownAction(LName, Lvalue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if (Object_Type.equalsIgnoreCase("captureStock"))
						{
							FunctionLibrary.captureStock(LName, Lvalue);
							logger.log(LogStatus.INFO, Description);

						}
						if (Object_Type.equalsIgnoreCase("stockTable"))
						{
							FunctionLibrary.stockTable();
							logger.log(LogStatus.INFO, Description);
						}
						if (Object_Type.equalsIgnoreCase("captureSupplier"))
						{
							FunctionLibrary.captureSupplier(LName,Lvalue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if (Object_Type.equalsIgnoreCase("supplierTable"))
						{
							FunctionLibrary.supplierTable(LName,Lvalue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if (Object_Type.equalsIgnoreCase("captureCustomer"))
						{
							FunctionLibrary.captureCustomer(LName,Lvalue, TestData);
							logger.log(LogStatus.INFO, Description);
						}
						if (Object_Type.equalsIgnoreCase("customerTable"))
						{
							FunctionLibrary.customerTable(LName,Lvalue, TestData);
							logger.log(LogStatus.INFO, Description);
						}

						//write as pass into TCModule sheet in status cell
						xl.setCellData(TCModule, j, 5, "Pass", outputpath);
						logger.log(LogStatus.PASS, Description);
						Module_Status="True";
					}catch(Exception e)
					{
						System.out.println(e.getMessage());
						//write as Fail into TCModule sheet in status cell
						xl.setCellData(TCModule, j, 5, "Fail", outputpath);
						logger.log(LogStatus.FAIL, Description);
						Module_Status ="False";
					}
					if(Module_Status.equalsIgnoreCase("True"))
					{
						//write as pass into TestCases Sheet
						xl.setCellData(TestCases, i, 3, "Pass", outputpath);
					}

					else
					{
						//write as Fail into TestCases Sheet
						xl.setCellData(TestCases, i, 3, "Fail", outputpath);
					}
					report.endTest(logger);
					report.flush();
				}
			}
			else
			{
				//write as blocked for test cases flag to N in TestCases Sheet
				xl.setCellData(TestCases, i, 3, "Blocked", outputpath);
			}
		}

	}
}
