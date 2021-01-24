package LoginTestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class BothCorrect 
{
	WebDriver driver;
	WritableWorkbook workbookcopy;
	WritableSheet writablesh;
	String testcase,result;
	int columnCount;

	/*String [][]Data = { 
			{"Admin","admin123"},
			{"Admin","Test"},
			{"Tester","admin123"},
			{"Tester","Test"}	};*/

	@BeforeSuite
	public void beforeSuite() throws BiffException, IOException,RowsExceededException,WriteException, InterruptedException
	{
		FileInputStream testData = new FileInputStream("C:\\Users\\naveenraj.durairaj\\Downloads\\testleaf.xls");
		Workbook workbook = Workbook.getWorkbook(testData);
		Sheet sheets = workbook.getSheet(0);
		int rowCount = sheets.getRows();
		int columnCount = sheets.getColumns();

		//Read rows and columns and save it in String Two dimensional array
		String inputdata[][] = new String[rowCount][columnCount];

		FileOutputStream testOutput = new FileOutputStream("C:\\Users\\naveenraj.durairaj\\Downloads\\testleafResults.xls");
		workbookcopy = Workbook.createWorkbook(testOutput);
		writablesh = workbookcopy.createSheet("results", 0);

		for(int i=0;i<rowCount;i++) 
		{
			for(int j=0;j<columnCount;j++) 
			{
				inputdata[i][j] =sheets.getCell(j, i).getContents();
				Label data = new Label(j, i, inputdata[i][j]);
				Label title = new Label(columnCount+1, 0, "Result");
				writablesh.addCell(data);
				writablesh.addCell(title);
			}
		}
	}


	@AfterSuite
	public void afterSuite() throws IOException,WriteException
	{
		workbookcopy.write();
		workbookcopy.close();
	}



	@DataProvider(name="LoginData") // starts with Upper Case D
	public String[][] LoginDataProvider() throws BiffException, IOException
	{
		String[][] Data = getExcelData();
		return Data;
	}

	public String[][]  getExcelData() throws BiffException, IOException
	{

		FileInputStream excel = new FileInputStream("C:\\Users\\naveenraj.durairaj\\Downloads\\testleaf.xls");
		Workbook workbook = Workbook.getWorkbook(excel);
		Sheet sheet = workbook.getSheet(0);

		int  rowCount =	sheet.getRows();
		int  colCount =	sheet.getColumns();
		System.out.println("Row value is "+rowCount);
		System.out.println("Column value is "+colCount);
		String testData[][] = new String[rowCount-1][colCount]; 
		//Data in string array will be stored as (0,0)=(row,column)
		//Data in Excel sheet will be stored as (0,0)=(column,row)

		for (int i=1;i<rowCount;i++) 
		{	for(int j=0;j<colCount;j++) 
		{
			testData[i-1][j] = sheet.getCell(j,i).getContents();
			System.out.println("Cell value is "+testData[i-1][j]);
		}
		}
		return testData;
	}



	//@Parameters({"UserName","Password"})

	@BeforeTest
	public void beforeTest()
	{
		System.setProperty("webdriver.ie.driver", 
				"C:\\MyCompetency_TestAutomation\\MyCompetency_Regression_Automation\\Servers\\IEDriverServer32.exe");

		WebDriver driver = new InternetExplorerDriver();
	}

	
	@Test(dataProvider = "LoginData") // starts with lower Case D
	public void LoginWithMultipleTestData(String UserName, String Password)
	{
		driver.get("https://opensource-demo.orangehrmlive.com/");

		WebElement userName = driver.findElement(By.id("txtUsername"));
		userName.sendKeys(UserName);

		WebElement password = driver.findElement(By.id("txtPassword"));
		password.sendKeys(Password);

		WebElement submit = driver.findElement(By.id("btnLogin"));
		submit.click();
		result = "true";
	}	

	@AfterMethod
	public void afterMethod() throws RowsExceededException,WriteException, InterruptedException
	{
		if(result.equalsIgnoreCase("true"))
		{
			testcase="PASS";

		}else
		{
			testcase = "FAIL";
		}
		Label report = new Label(columnCount+1, 4, testcase);
		writablesh.addCell(report);
		
		
	}



	@AfterTest
	public void afterTest()
	{
		driver.quit();
	}



}
