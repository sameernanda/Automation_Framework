package com.nexenta.ftaf.glue.S417.fusionGUI;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.Select;

import com.nexenta.ftaf.glue.S417.S417;
import com.nexenta.ftaf.utilities.FtafUtilities;
import com.nexenta.ftaf.utilities.connectors.testrail.entities.TestStepResult;
import com.nexenta.ftaf.utilities.constants.FtafConstants.TestResultStatus;
import com.nexenta.ftaf.utilities.parsers.yaml.YamlParser;
public class FusionGuiConnector extends S417{
	WebDriver driver;
	List<WebElement> dropdown;
	String browserName;
	static String returnMsg;
	static String actualOutput;
	static String _currentStep="";
	static String dateAct = new SimpleDateFormat("yyyyMMddhhmm").format(new Date());
	public FusionGuiConnector() {
	}
	public String open_Browser(String browser){
		browserName = browser;
		switch(browser){
		case "chrome" :
			System.setProperty("webdriver.chrome.driver","E:\\Automation_Framework_v1\\Driver\\chromedriver_win32\\chromedriver.exe");
			driver = new ChromeDriver();
			break; 
		case "firefox" :
			driver = new FirefoxDriver();
			driver.manage().window().maximize();
			break;
		case "IE" :
			System.setProperty("webdriver.ie.driver","E:\\Automation_Framework_v1\\Driver\\IEDriverServer_x64_2.45.0\\IEDriverServer.exe");
			driver = new InternetExplorerDriver();
			break;
		}
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		return browser+" is successfully opened";
	}

	private String open_URL(Cell urlPath){
		try{
			if(browserName.equalsIgnoreCase("IE")){
				driver.get(urlPath.toString());
				driver.get("javascript:document.getElementById('overridelink').click();");
			}else{
				driver.get(urlPath.toString());
			}
			returnMsg = "Successfully opened "+urlPath;
			_testStepResults.add(new TestStepResult(_currentStep,returnMsg, TestResultStatus.PASSED.getStatusCode()));
		}catch(Exception pe){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Error occured while opening Url \"+urlPath+\" because \"+pe.getMessage()"), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException("Error occured while opening Url "+urlPath+" because "+pe.getMessage());
		}
		return returnMsg;
	}

	private String enter_Text(Cell property,Cell propertyValue,Cell input){
		try{
			if(property.toString().equals("id")){
				driver.findElement(By.id(propertyValue.toString())).sendKeys(input.toString());
			}else if(property.toString().equals("xpath")){
				driver.findElement(By.xpath(propertyValue.toString())).sendKeys(input.toString());
			}
			returnMsg = "Successfully entered input in "+property+" " +propertyValue;
			_testStepResults.add(new TestStepResult(_currentStep,returnMsg, TestResultStatus.PASSED.getStatusCode()));
		}catch(Exception pe){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Error occured while writing input to "+property+" because "+pe.getMessage()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException("Error occured while writing input to "+property+" because "+pe.getMessage());
		}
		return returnMsg;
	}

	private String click_Link(Cell property,Cell propertyValue){
		try{
			if(property.toString().equals("id")){
				driver.findElement(By.id(propertyValue.toString())).click();
			}else if(property.toString().equals("xpath")){
				driver.findElement(By.xpath(propertyValue.toString())).click();
			}returnMsg = "Successfully clicked on "+propertyValue;
			_testStepResults.add(new TestStepResult(_currentStep,returnMsg, TestResultStatus.PASSED.getStatusCode()));
		}catch(Exception pe){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Error occured while clicking on the element "+property+" because "+pe.getMessage()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException("Error occured while clicking on the element "+property+" because "+pe.getMessage());
		}
		return returnMsg;
	}

	private String check_Current_URL(){
		try{
			returnMsg = "Current URL is  : "+driver.getCurrentUrl();
			System.out.println("Current TITLE of browser is  : "+driver.getTitle());
			_testStepResults.add(new TestStepResult(_currentStep,returnMsg, TestResultStatus.PASSED.getStatusCode()));
		}catch(Exception pe){
			throw new RuntimeException("Error occured while checking current URL because "+pe.getMessage());
		}
		return returnMsg;
	}

	private void quit_Browser() throws InterruptedException{
		if(driver!=null){
			Thread.sleep(3000);
			driver.quit();
		}
	}

	public String getText(Cell property, Cell propertyValue) {
		try{
			if(property.toString().equals("id")){
				actualOutput = driver.findElement(By.id(propertyValue.toString())).getText();
			}else if(property.toString().equals("xpath")){
				actualOutput = driver.findElement(By.xpath(propertyValue.toString())).getText();
			}
			_testStepResults.add(new TestStepResult(_currentStep,"Successfully retrived text: "+actualOutput, TestResultStatus.PASSED.getStatusCode()));

		}catch(RuntimeException pe){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Error occurred while getting Attribute of "+propertyValue+ pe.getMessage()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException("Error occurred while getting Attribute of "+propertyValue+ pe.getMessage());
		}
		return actualOutput;
	}

	public String verifyText(Cell property, Cell propertyValue,Cell expectedOutput) {
		actualOutput = this.getText(property, propertyValue);
		if(!actualOutput.equalsIgnoreCase(expectedOutput.toString())){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Actual output: \"%s\" \n does not match with \n Expected output: \"%s\"",actualOutput,expectedOutput.toString()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException(String.format("Actual output: \"%s\" does not match with Expected output: \"%s\"",actualOutput,expectedOutput.toString() ));
		}else{
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Actual output: \"%s\" \n matched with \n Expected output: \"%s\"",actualOutput,expectedOutput.toString()), TestResultStatus.PASSED.getStatusCode()));
			return (String.format("Actual output: \"%s\" match with Expected output: \"%s\"",actualOutput,expectedOutput.toString() ));
		}
	}

	public List<WebElement> getmultipleWebElements(Cell property, Cell propertyValue) {
		try{
			if(property.toString().equals("id")){
				dropdown = driver.findElements(By.id(propertyValue.toString()));
			}else if(property.toString().equals("xpath")){
				dropdown = driver.findElements(By.xpath(propertyValue.toString()));
			}
			_testStepResults.add(new TestStepResult(_currentStep,"Successfully retrived dropdown values: "+printGetText(dropdown), TestResultStatus.PASSED.getStatusCode()));
		}catch(RuntimeException pe){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Error occurred while getting dropdown values of "+propertyValue+ pe.getMessage()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException("Error occurred while getting dropdown values of "+propertyValue+ pe.getMessage());
		}
		return dropdown;
	}

	public void verifyWebElementPresent(Cell property, Cell propertyValue, Cell input) throws InterruptedException {
		Thread.sleep(6000);
		String actualOutput = null;
		dropdown = this.getmultipleWebElements(property, propertyValue);
		boolean flag = false;
		for(WebElement web: dropdown){
			if(web.getText().equals(input.toString())){
				flag = true;
			}
			actualOutput = web.getText();
		}
		if (!flag == true){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Actual output: \"%s\" \n does not match with \n Expected output: \"%s\"",actualOutput,input.toString()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException(String.format("Actual output: \"%s\" does not match with Expected output: \"%s\"",actualOutput,input.toString() ));
		}else{
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Actual output: \"%s\" \n matched with \n Expected output: \"%s\"",actualOutput,input.toString()), TestResultStatus.PASSED.getStatusCode()));
		}
	}	

	private StringBuilder printGetText(List<WebElement> allOptions) {
		StringBuilder stringBuilderObjectToReturn=new StringBuilder();
		for(WebElement web: allOptions){
			stringBuilderObjectToReturn.append("\n"+web.getText());
		}
		return stringBuilderObjectToReturn;
	}
	public String getAttribute(Cell property, Cell propertyValue, Cell input) {
		try{
			if(property.toString().equals("id")){
				actualOutput = driver.findElement(By.id(propertyValue.toString())).getAttribute(input.toString());
			}else if(property.toString().equals("xpath")){
				actualOutput = driver.findElement(By.xpath(propertyValue.toString())).getAttribute(input.toString());
			}
			_testStepResults.add(new TestStepResult(_currentStep,"Successfully retrived Attribute "+input+" as : "+actualOutput, TestResultStatus.PASSED.getStatusCode()));
		}catch(RuntimeException pe){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Error occurred while getting Attribute of "+propertyValue+ pe.getMessage()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException("Error occurred while getting Attribute of "+propertyValue+ pe.getMessage());
		}
		return actualOutput;
	}

	public String verifyAttribute(Cell property, Cell propertyValue,Cell input,Cell expectedOutput) {
		actualOutput = this.getAttribute(property, propertyValue,input);
		if(!actualOutput.equalsIgnoreCase(expectedOutput.toString())){
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Actual output: \"%s\" \n does not match with \n Expected output: \"%s\"",actualOutput,expectedOutput.toString()), TestResultStatus.FAILED.getStatusCode()));
			throw new RuntimeException(String.format("Actual output: \"%s\" does not match with Expected output: \"%s\"",actualOutput,expectedOutput.toString() ));
		}else{
			_testStepResults.add(new TestStepResult(_currentStep,String.format("Actual output: \"%s\" \n matched with \n Expected output: \"%s\"",actualOutput,expectedOutput.toString()), TestResultStatus.PASSED.getStatusCode()));
			return (String.format("Actual output: \"%s\" match with Expected output: \"%s\"",actualOutput,expectedOutput.toString() ));
		}
	}

	protected static final String _testRunID = ((Map<String, String>) YamlParser.load(_contextLocation + "TestRail.yaml")).get("testRunID");
	protected static final String _applianceVersion = _applianceInfo.get(_platform).get("Appliance").get("applianceVersion");
	protected static final String _testRailplatformValue = String.valueOf(_applianceInfo.get(_platform).get("Platform").get("appliancePlatform"));

	public static void main(String[] args) throws InvalidFormatException, IOException, InterruptedException, FileNotFoundException   {
		FusionGuiConnector instance = new FusionGuiConnector();
		//Input Flat File
		FileInputStream flatFile = new FileInputStream("E:/Automation_Framework_v1/Driver/Driver.xls");
		Workbook flatFileWorkbk = WorkbookFactory.create(flatFile);
		Sheet sheetToStart= flatFileWorkbk.getSheet("Sheet1");
		for(int m=1;m<=sheetToStart.getLastRowNum();m++){
			Cell testCaseID = sheetToStart.getRow(m).getCell(0);
			Cell browser = sheetToStart.getRow(m).getCell(2);
			Cell execute = sheetToStart.getRow(m).getCell(3);
			Cell iteration = sheetToStart.getRow(m).getCell(4);
			int noOfIteration = Double.valueOf(iteration.toString()).intValue();
			//Input Data File
			String inputFile = null;
			if(testCaseID.toString()!="" && !testCaseID.toString().equals(null)){
				inputFile = "E:/Automation_Framework_v1/TestCases/"+testCaseID+".xls";
			}
			FileInputStream fis = new FileInputStream(inputFile);
			Workbook wb = WorkbookFactory.create(fis);
			Sheet inputSheet= wb.getSheet("Sheet1");
			if(execute.toString().equalsIgnoreCase("Yes")){
				try{
					while(noOfIteration > 0){
						instance.open_Browser(browser.toString());
						for(int i =1;i<=inputSheet.getLastRowNum();i++){
							Cell action =  inputSheet.getRow(i).getCell(2);
							if(!(action == null)){
								switch (action.toString()){
								case("open_URL"):
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
								instance.open_URL(inputSheet.getRow(i).getCell(5));
								break;
								case("enter_Text"):
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
								instance.enter_Text(inputSheet.getRow(i).getCell(3), inputSheet.getRow(i).getCell(4), inputSheet.getRow(i).getCell(5));
								break;
								case "click_Link" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									Thread.sleep(3000);
									instance.click_Link(inputSheet.getRow(i).getCell(3), inputSheet.getRow(i).getCell(4));
									break;
								case "check_Current_URL" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									instance.check_Current_URL();
									break;
								case "getText" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									Thread.sleep(3000);
									instance.getText(inputSheet.getRow(i).getCell(3), inputSheet.getRow(i).getCell(4));
									break;
								case "getAttribute" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									Thread.sleep(3000);
									instance.getAttribute(inputSheet.getRow(i).getCell(3),inputSheet.getRow(i).getCell(4),inputSheet.getRow(i).getCell(5));
									break;
								case "verifyText" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									Thread.sleep(3000);
									instance.verifyText(inputSheet.getRow(i).getCell(3),inputSheet.getRow(i).getCell(4), inputSheet.getRow(i).getCell(6));
									break;	
								case "verifyAttribute" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									Thread.sleep(3000);
									instance.verifyAttribute(inputSheet.getRow(i).getCell(3),inputSheet.getRow(i).getCell(4),inputSheet.getRow(i).getCell(5), inputSheet.getRow(i).getCell(6));
									break;	
								case "getmultipleWebElements" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									Thread.sleep(6000);
									instance.getmultipleWebElements(inputSheet.getRow(i).getCell(3), inputSheet.getRow(i).getCell(4));
									break;
								case "verifyWebElementPresent" :
									_currentStep=inputSheet.getRow(i).getCell(1).toString();
									Thread.sleep(3000);
									instance.verifyWebElementPresent(inputSheet.getRow(i).getCell(3), inputSheet.getRow(i).getCell(4),inputSheet.getRow(i).getCell(6));
									break;
								}
							}
						}noOfIteration --;
					}
					instance.quit_Browser();
				}catch(Exception e){
					instance.quit_Browser();
					throw e;	
				}
				finally{
					publishTestStepResultsForFusionGUI(inputSheet.getRow(1).getCell(0).toString(),_testRunID, _applianceVersion, FtafUtilities.getTestRailCompatableTimeFormat(_startTime, _endTime),_testRailplatformValue);
				}
			}
		}
	}
}