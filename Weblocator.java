package com.util;
import java.awt.AWTException;
import java.awt.Robot;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.sql.Connection;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.Random;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;


public class Weblocator extends Setup {
	//public static By by;
	static String destDir;
	public static boolean browserAlreadyOpen=false;
	public static Properties SYSPARAM =null;
	//public static Logger Add_Log = null;
	public static Properties Object = null;
	public static Connection conn;
	static String runningTest1 = "error geting";
	Dimension size;
	public static String output;

	public static String filelocation;
	public static  FileInputStream ipstr = null;
	public  static FileOutputStream opstr =null;
	public static int value;
	//---------------css---------------
	public static String fontColor=null;

	//--------------------excle-------------

	//-------------------jxl---------------------
	public static int totalNoOfRows; 
	public static int totalNoOfCols;
	public static Sheet sh;
	static int i;
	public static WritableWorkbook myFirstWbook;
	public static WritableSheet excelSheet;


	public static void random() {
		Random rand = new Random(); 
		value = rand.nextInt(900);
		explicitWait(1);
	}

	public static boolean verify_Element_Is_Enabled(By link){
		boolean isEnable=false;
		WebElement firstname = waitForOptionalElement(link);
		if (firstname != null && firstname.isDisplayed()) {
			Log.pass("WebElement List items: "+firstname.getText()+" Displayed and check Element is ENABLE");
			firstname.isEnabled();
			Log.pass("WebElement "+firstname.getText()+" is ENABLE");
			isEnable=true;
		} 
		return isEnable;
	}

	public static WebElement waitForOptionalElement(By locator) {
		WebElement element = null;
		int count = 0;
		int maxTries = 10;
		while (true) {
			try {
				Wait<WebDriver> wait= new FluentWait<WebDriver>(driver).withTimeout(40, TimeUnit.SECONDS).pollingEvery(6,  TimeUnit.SECONDS).ignoring(NoSuchElementException.class);
				List<WebElement> elementList = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(locator));
				if (elementList.size() != 0) {
					element = elementList.get(0);
				}
				return element;
			} catch (StaleElementReferenceException e) {
				explicitWait(3);
				Log.info("Stale element occurs in geting the element while waiting, retry value: " + count + " :");
				if (++count == maxTries) {
					throw e;
				}
			} catch (TimeoutException ex) {
				Log.error("Optional element is not found:"+ex.getMessage());
				return element;
			}
		}
	}

	public static boolean RadioBtn(By link)  {
		boolean text = false;
		WebElement textpresent = waitForOptionalElement(link);
		if (textpresent != null && textpresent.isDisplayed()) {
			//Check if radio button is selected by default
			Log.pass( "WebElement "+textpresent.getText()+" Displayed and Select Radio Btn");
			if (textpresent.isSelected()) {
				// Print message to console
				System.out.println(" radio button is selected by default");
				text = true;
			} else {
				// Click the radio button
				System.out.println("Radio button is not selected");
			}
		} 
		return text;
	}

	public static String GetDateTime() {
		DateFormat dateFormat = new SimpleDateFormat("ddMMMyy_hh:mm:ssaa");
		Date date = new Date();
		String destFile = dateFormat.format(date);
		Log.pass("Get Current Date and Time: "+destFile);
		System.out.println("Get current Date & Time:  "+destFile);
		return destFile;
	}

	public static boolean clearText(By link)  {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			//linkpresent.click();
			Log.pass("WebElement List items: "+linkpresent.getText()+" Displayed and clear text Field");
			//linkpresent.clear();
			for (int i = 0; i <=8; i++) {
				linkpresent.sendKeys(Keys.BACK_SPACE);
			}
			links = true;

		} 
		return links;
	}


	public static boolean PressEnterBtn(By link) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement "+linkpresent.getText()+" Displayed and Press Enter Btn");

			linkpresent.sendKeys(Keys.ENTER);
		} 
		return links;
	}

	public static void Reload() {
		Log.pass("Refresh and Reload");
		driver.navigate().refresh();
		try {
			driver.switchTo().alert().accept();
		}catch (Exception e) {
			System.out.println("Alert Not present");
		}
	}

	public static boolean Openlinks(By locator)  {
		//explicitWait(1);
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(locator);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement "+linkpresent.getText()+" Displayed and click.");
			linkpresent.click();
			//explicitWait(1);
			links = true;

		} 
		return links;
	}

	public static boolean DoubleClick(By link) {
		Actions action = new Actions(driver);
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement "+linkpresent.getText()+" Displayed and Double Click ");
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", linkpresent);
			action.doubleClick(linkpresent).build().perform();
			links = true;
		} 
		return links;
	}

	public static WebElement webelement(By link){
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()){
			Log.pass("WebElement "+element.getText()+" Displayed");
		}
		return element;
	}

	public static WebElement webElementList(By link){
		List<WebElement> elemtList = null;
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()){
			//Log.pass("WebElement "+element.getText()+" Displayed");
			elemtList = element.findElements(link);
		}
		return (WebElement) elemtList;
	}

	/*	public static boolean TextField_new(By link, String ... entertext) {  // By Deepak
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			links = true;
			//linkpresent.clear();
			linkpresent.click();
			try {
				linkpresent.sendKeys(entertext[0]);
				if(entertext.length==2) {
					System.out.println(entertext[1]);
				}
				try {
					TimeUnit.SECONDS.sleep(1);
					driver.findElement(link).sendKeys(Keys.TAB);
					TimeUnit.MILLISECONDS.sleep(1);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			System.out.println(entertext[0]+" :  Enter data");
		} else {
			System.out.println(entertext[0]+" :  not enter on "+link);
		}
		return links;
	}*/

	public static boolean TextField(By link, String entertext) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			links = true;
			linkpresent.clear();
			linkpresent.click();
			Log.pass("WebElement "+linkpresent.getText()+" Displayed and Enter Text :"+entertext+" And Click on TAB Button");
			linkpresent.sendKeys(entertext);
			//explicitWait(1);
			driver.findElement(link).sendKeys(Keys.TAB);
			explicitWait(1);
		}
		return links;
	}

	public static boolean clickOnTAB(By link) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement "+linkpresent.getText()+" Displayed and Click on TAB Button");
			driver.findElement(link).sendKeys(Keys.TAB);
		}
		return links;
	}

	public static boolean TextFieldWithOutTAB(By link, String entertext) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			links = true;
			Log.pass("WebElement "+linkpresent.getText()+" Displayed and Enter Text :"+entertext);
			linkpresent.sendKeys(entertext);
			System.out.println(entertext+" : is Enter data");
		} 
		return links;
	}

	public static boolean TextFieldJavascriptExecutor(By link, String entertext) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement "+linkpresent.getText()+" Displayed and Enter Text :"+entertext);
			JavascriptExecutor jse = (JavascriptExecutor)driver;
			jse.executeScript("arguments[0].value="+"'"+entertext+"'"+";", linkpresent);
			links = true;
			clickOnTAB(link);
			/*	try {
				driver.hideKeyboard();// for Mobile
			} catch (Exception e) {
				//logger.info("hideKeyboard()");
			}*/
			////logger.info(linkpresent.getText()+" : is Enter data");
		}
		return links;
	}

	public static boolean TextField_Clean(By link) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement "+linkpresent.getText()+" Displayed.");
			links = true;
			linkpresent.clear();
		}
		return links;
	}

	public static boolean Syncing_BufferingContent(By link)    // Wait finction
	{
		System.out.println("user enter the loading section");
		Boolean message=false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement "+linkpresent.getText()+" Displayed.");
			while(driver.findElements(link).size()!= 0){}
			System.out.println("user out the loading section");
		}
		return message;
	}

	public static boolean IselementPresent(By textvalue)  {
		boolean text = false;
		WebElement textpresent = waitForOptionalElement(textvalue);
		if (textpresent != null && textpresent.isDisplayed()) {
			Log.pass("WebElement "+textpresent.getText()+" Present.");
			text = true;
		} 
		return text;
	}

	public static String getPagetext(By gettext) {
		//WebDriverManager.explicitWait(1);
		String Actualtext = "";
		WebElement element = waitForOptionalElement(gettext);
		if (element != null && element.isDisplayed()) {
			Log.pass("WebElement "+element.getText()+" Displayed. and Get Actual Text.");
			Actualtext = element.getText();
			//logger.info("Get Text:  " + Actualtext);
		}
		return Actualtext;
	}

	public static String Getactualtext(By gettext) {
		String Actualtext = "";
		WebElement element = waitForOptionalElement(gettext);
		if (element != null && element.isDisplayed()) {
			Actualtext = element.getText();
			Log.pass("WebElement "+element.getText()+" Displayed. and Get Actual Text.");
		}
		return Actualtext;
	}

	public static String GetAttributevalue(By link) {
		String Actualtext = "";
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()) {
			Log.pass("WebElement "+element.getAttribute("value")+" Displayed. and Get Attribute Text.");
			Actualtext = element.getAttribute("value");
		}
		return Actualtext;
	}

	public static void getListofLink(By link) {
		explicitWait(1);
		System.out.println("To check and get share by links in the phone.");
		WebElement textpresent = waitForOptionalElement(link);
		if (textpresent != null && textpresent.isDisplayed()) {
			Log.pass("WebElement List items: "+textpresent.getText()+" Displayed and Get List");
			List<WebElement> allSuggestions = driver.findElements(link);
			for (WebElement suggestion : allSuggestions) {
				i++;
				//System.out.println("List By Items:   \n"+i+"). "+ suggestion.getText());
				Log.pass("WebElement List items: "+suggestion.getText());
			}
		} 
	}

	public static void keyPressALTQ(By link) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement List items: "+linkpresent.getText()+" Displayed and keyPress: ALT + Q ");
			linkpresent.click();
			linkpresent.sendKeys(Keys.chord(Keys.ALT, "q"));
		}
	}

	public static Boolean genralkeyPress(int keyNo) throws AWTException {
		boolean status = false;
		try {
			Log.pass("General Press KEYS"+keyNo);
			Robot robot=new Robot();
			robot.keyPress(keyNo);
			explicitWait(1);
			status = true;
			//robot.keyPress(KeyEvent.VK_S);
			//robot.keyRelease(KeyEvent.VK_ALT);
			// robot.keyRelease(KeyEvent.VK_S);        
		}catch (Exception e) {

		}

		return status;
	}

	public static Boolean genralkeyPressUsingAction(int keyNo) throws AWTException {
		boolean status = false;
		try {
			Log.pass("General Press KEYS"+keyNo);

			Actions action = new Actions(driver);
			//action.sendKeys(Keys.F12);
			action.sendKeys("Keys."+keyNo);
			action.perform();
			status = true;

		}catch (Exception e) {

		}

		return status;
	}


	public static void getWindowHandle() {
		explicitWait(2);
		Log.pass("Newly open Window/Frame Handle.");
		driver.getWindowHandle(); 
		explicitWait(1);
	}

	public static void explicitWait(int timeSeconds) {
		try {
			TimeUnit.SECONDS.sleep(timeSeconds);
		} catch (InterruptedException e) {
			Log.warn("Error in TimeUnit wait");
		}
	}

	public static void scrollingByCoordinatesofAPage() {  // top to buttom
		Log.pass("Scroll By Coordinate");
		((JavascriptExecutor) driver).executeScript("window.scrollBy(0,500)");
		explicitWait(2);
	}

	public static void scrollingByCoordinatesofAPage_scrollUp() {  // buttom  to top
		Log.pass("ScrollUp Page ");
		((JavascriptExecutor) driver).executeScript("window.scrollBy(0,-500)");
		explicitWait(2);
	}

	public static String printExceptionTrace(Exception e) {
		StringWriter errors = new StringWriter();
		e.printStackTrace(new PrintWriter(errors));
		Log.pass("Calling Print Exception Trace some thing Error is present");
		Log.error("Error : "+errors.toString());
		return errors.toString();
	}

	public static void scrollingToBottomofAPage(By link) {
		Log.pass("Scroll To Bottom page");
		((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
	}

	public static void mouse(By link) {
		Actions action = new Actions(driver);
		WebElement we = driver.findElement(link);
		action.moveToElement(we).build().perform();
	}

	public static String randomestring()
	{
		Log.pass("Call RandomString method");
		String generatedstring=RandomStringUtils.randomAlphabetic(8);
		return(generatedstring);
	}

	public static String randomeNum() {
		Log.pass("Call RandomNumber method");
		String generatedString2 = RandomStringUtils.randomNumeric(5);
		return (generatedString2);
	}


	public static boolean TwoValueReturn(By link1, By link2){
		int count=0;
		boolean text = false;

		List<WebElement> Getlist1 =  driver.findElements(link1);    
		List<WebElement> Getlist2 =  driver.findElements(link2);  

		if(Getlist1!=null){
			text=true;
			for(int i=0;i<Getlist1.size();i++){
				count++;
				System.out.println(count+"). "+Getlist1.get(i).getText()+">>> "+Getlist2.get(i).getText());
			}
		}
		return text;

	}

	public static void listvalueclick(By link,String value){
		List<WebElement> list = driver.findElements(link);
		for (WebElement element : list) 
		{ 
			if(element.getAttribute("value").equals(value)) 
			{ 
				element.click(); 
			} 
		}
	}

	public static boolean CheckingChkbox(WebElement chkbx1){
		Log.pass("Checking Check Box");
		boolean checkstatus=false;
		checkstatus=chkbx1.isSelected();
		if (checkstatus==true){
			Log.pass("Check Box already checked");
			checkstatus=true;
		}
		else
		{
			chkbx1.click();
			System.out.println("Checked the checkbox");
		}
		return checkstatus;
	}

	public static boolean OpenlinksJavascriptExecutor(By link)  {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			JavascriptExecutor executor = (JavascriptExecutor)driver;
			executor.executeScript("arguments[0].click();", linkpresent);
			links = true;
		} 
		return links;
	}

	public static File takeScreenshot(String filename) throws IOException {
		Log.pass("Take a ScreenShot");
		//destDir = SYSPARAM.getProperty("screenshotpath");
		destDir = System.getProperty("user.dir")+"//Screenshot";
		//driver.context("NATIVE_APP");
		File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		new File(destDir).mkdirs();
		String destFile = filename + ".png";
		FileUtils.copyFile(scrFile, new File(destDir + "//" + destFile));
		//PageElement.changeContextToWebView(driver);
		return scrFile; 
	}

	public static boolean mouse_hover_event(By linkMenu, By menuOption){
		Log.pass("Call Mouse Hover Method");
		boolean hoveringMouse=false;
		Actions actions = new Actions(driver);
		WebElement moveonmenu = waitForOptionalElement(linkMenu);
		if (moveonmenu != null && moveonmenu.isDisplayed()) {
			actions.moveToElement(moveonmenu);
			hoveringMouse=true;
			WebElement hoverOption = waitForOptionalElement(menuOption);
			if (hoverOption != null && hoverOption.isDisplayed()) {
				actions.moveToElement(driver.findElement(menuOption)).click().perform();

			}
			else{
				System.out.println("Hover to menu option is not present");
			}
		}
		return hoveringMouse;

		/*Actions actions = new Actions(driver);
		WebElement moveonmenu = driver.findElement(linkMenu);
		actions.moveToElement(moveonmenu).moveToElement(driver.findElement(menuOption)).click().perform();*/
	}

	public static void window_handles(String window1_or_window2){
		Log.pass("Multiple Window Handle");

		Set<String> AllWindowHandles = driver.getWindowHandles();
		String window1 = (String) AllWindowHandles.toArray()[0];
		System.out.print("window1 handle code = "+AllWindowHandles.toArray()[0]);
		String window2 = (String) AllWindowHandles.toArray()[1];
		System.out.print("\nwindow2 handle code = "+AllWindowHandles.toArray()[1]);
		//Switch to window2(child window) and performing actions on it.
		driver.switchTo().window(window1_or_window2);
	}


	public static String AlertAcceptWithGetMsg() {
		String msg = null;
		Log.pass("Alert Accept with Get Message");
		try {
			Alert alert=driver.switchTo().alert();
			msg=alert.getText();
			alert.accept();
		}catch (Exception e) {
			System.out.println("Alert Not present");
		}
		return msg;
	}

	public static void AlertAccept() {
		Log.pass("Alert Accept");
		try {
			driver.switchTo().alert().accept();
		}catch (Exception e) {
			System.out.println("Alert Not present");
		}
	}

	public static void SwitchToAlert() {
		Log.pass("Switch To Alert");
		try {
			driver.switchTo().alert();
		}catch (Exception e) {
			System.out.println("Alert Not present");
		}
	}

	public static void handle_unexpected_alert(){
		Log.pass("Handle Unexpected Alert method");
		try{   
			driver.switchTo().alert().dismiss();  
		}catch(Exception e){ 
			System.out.println("unexpected alert not present");   
		}
	}

	public static void highlight_Element(By link){
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()) {
			JavascriptExecutor js = (JavascriptExecutor)driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');", element);
			// js.executeScript("arguments[0].style.border='4px groove green'", element);
		}
	}

	public static void css_Reading_Font_Properties_reading(By link){
		//Locate text string element to read It's font properties.
		WebElement text = driver.findElement(link);
		System.out.println("--------------------------------------------------"); 
		//Read font-size property and print It In console.
		String fontSize = text.getCssValue("font-size");
		System.out.println("Font Size -> "+fontSize);

		//Read color property and print It In console.
		String fontColor = text.getCssValue("color");
		System.out.println("Font Color -> "+fontColor);

		//Read font-family property and print It In console.
		String fontFamily = text.getCssValue("font-family");
		System.out.println("Font Family -> "+fontFamily);

		//Read text-align property and print It In console.
		String fonttxtAlign = text.getCssValue("text-align");
		System.out.println("Font Text Alignment -> "+fonttxtAlign);
		System.out.println("--------------------------------------------------");
	}

	public static long DifferentbtweenStartEndTime(){
		Log.pass("Get Differece Btween Start and End Time");
		long start_time=GetStartTime();
		//Get the difference (currentTime - startTime)  of times.		
		System.out.println("Passed time: " + (System.currentTimeMillis() - start_time));
		long AvgTime=(System.currentTimeMillis() - start_time);
		return AvgTime;

	}

	public static long GetStartTime(){
		Log.pass("Get Start Time");
		//Declare and set the start time		
		long start_time = System.currentTimeMillis();
		return start_time;	

	}

	public static void BrowserAndOSDetails() {
		Log.pass("Get Brower Details");
		try {
			//Get Browser name and version.
			/*Capabilities caps = ((RemoteWebDriver) driver).getCapabilities();
			String browserName = caps.getBrowserName();
			String browserVersion = caps.getVersion();

			//Get OS name.
			String os = System.getProperty("os.name").toLowerCase();
			try {
				Log.pass("OS = " + os + ", Browser = " + browserName + " "+ browserVersion);
			}
			catch (Exception e) {
			}
			System.out.println("OS = " + os + ", Browser = " + browserName + " "+ browserVersion);*/

			Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
			String browserName = cap.getBrowserName().toLowerCase();
			System.out.println(browserName);
			String os = cap.getPlatform().toString();
			//System.out.println(os);
			String v = cap.getVersion().toString();
			// System.out.println(v);
			Log.pass("OS = " + os + ", Browser = " + browserName + " "+ v);

		}catch (Exception e) {

		}

	} 

	//-----------------------Excle--------------------------
	public static void ReadExcel() throws BiffException, IOException {
		Setup setup=new Setup();
		String FilePath =setup.excelPath;
		String sheetName=setup.excel_sheetname;
		FileInputStream fs = new FileInputStream(FilePath);
		Workbook wb = Workbook.getWorkbook(fs);
		// TO get the access to the sheet
		//sh = wb.getSheet(sheetno);
		sh = wb.getSheet(sheetName);
		// To get the number of rows present in sheet
		totalNoOfRows = sh.getRows();
		System.out.println("Total No of Row:  "+totalNoOfRows);

		// To get the number of columns present in sheet
		totalNoOfCols = sh.getColumns();
		System.out.println("Total No of Coloum:  "+totalNoOfCols);
	}

	public static void WriteExcel() throws BiffException, IOException {
		Setup setup=new Setup();
		String FilePath =setup.excelPath;
		String sheetName=setup.excel_sheetname;
		//1. Create an Excel file
		myFirstWbook = null;
		Workbook workbook1 = Workbook.getWorkbook(new File(FilePath));
		myFirstWbook = Workbook.createWorkbook(new File(FilePath),workbook1);

		// create an Excel sheet
		// WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 2", 3); // if you want to create new sheet
		//excelSheet = myFirstWbook.getSheet(sheetno);  //exciting sheet updating
		excelSheet = myFirstWbook.getSheet(sheetName);

		/*
			// add something into the Excel sheet
			Label label = new Label(colno, rowno, writedata);
			excelSheet.addCell(label);

			Number number = new Number(0, 1, 1);
			excelSheet.addCell(number);

			label = new Label(1, 0, "Result");
			excelSheet.addCell(label);

			label = new Label(1, 1, "Passed");
			excelSheet.addCell(label);

			number = new Number(0, 2, 2);
			excelSheet.addCell(number);

			label = new Label(1, 2, "Passed 2");
			excelSheet.addCell(label);
			myFirstWbook.write();
			myFirstWbook.close();
		 */
	}

	public static void WriteTextFile() throws IOException {
		String text = "";

		String TestFile = "/Users/A16P6EEX/eclipseTestAutomation/RetailerPortal_Automation_Selenium/OutputFolder/temp.text";

		File FC = new File(TestFile);//Created object of java File class.

		try {
			FC.createNewFile();
		} catch (IOException e2) {
			e2.printStackTrace();
		}

		FileWriter FW = new FileWriter(TestFile);

		BufferedWriter BW = new BufferedWriter(FW);

		BW.write(text); //Writing In To File.

		BW.newLine();//To write next string on new line.

		//BW.flush();  /* ye line if you want to create CSV file vena delete or comment kr do or uppar bhi  file ka extance bhi change kr do */

		BW.close();

		//


	}
	public static boolean TextField_Clear_Click_Tab(By link, String entertext) {
		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			links = true;
			linkpresent.clear();
			linkpresent.click();

			linkpresent.sendKeys(entertext);
			explicitWait(1);
			driver.findElement(link).sendKeys(Keys.TAB);
			explicitWait(1);
			System.out.println(entertext+" :  Enter data");

		}
		return links;
	}

	public static boolean Openlinks_Click_Tab(By link)  {

		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			linkpresent.click();
			links = true;
		} 
		return links;
	}

	public static boolean drop_down_by_value(By link , String value){
		boolean select=false;
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()) {
			//Selecting value from drop down using visible text
			Select mydrpdwn = new Select(driver.findElement(link));
			WebElement element1 = waitForOptionalElement(link);
			if (element1 != null && element1.isDisplayed()) {
				Log.pass("WebElement List items: "+element.getText()+" Displayed");
				mydrpdwn.selectByVisibleText(value);
			}
			select=true;
		}
		return select;
	}

	public static boolean drop_down_by_index(By link, int indexno){
		boolean dropdownindex=false;
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()) {
			//Selecting value from drop down by index
			Select listbox = new Select(driver.findElement(link));
			WebElement element1 = waitForOptionalElement(link);
			if (element1 != null && element1.isDisplayed()) {
				Log.pass("WebElement List items: "+element.getText()+" Displayed");
				listbox.selectByIndex(indexno); // select index no
			}
			dropdownindex=true;
		}
		return dropdownindex;
	}

	public static boolean dropdown_using_visible_text(By link, String value){
		boolean select=false;
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()) {
			//Selecting value from drop down using visible text
			Select mydrpdwn = new Select(driver.findElement(link));
			WebElement element1 = waitForOptionalElement(link);
			if (element1 != null && element1.isDisplayed()) {
				Log.pass("WebElement List items: "+element.getText()+" Displayed");
				mydrpdwn.selectByVisibleText(value);
			}
			select=true;
		}
		return select;
	}

	public static boolean dropdown_deSelecting_using_visible_text(By link, String value){
		boolean deselect=false;
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()) {
			//Selecting value from drop down using visible text
			Select mydrpdwn = new Select(driver.findElement(link));

			WebElement element1 = waitForOptionalElement(link);
			if (element1 != null && element1.isDisplayed()) {
				Log.pass("WebElement List items: "+element.getText()+" Displayed");
				mydrpdwn.deselectByVisibleText(value);
			}
			deselect=true;
		}
		return deselect;
	}

	public static boolean uploadFile(By link, String filePath){

		boolean links = false;
		WebElement linkpresent = waitForOptionalElement(link);
		if (linkpresent != null && linkpresent.isDisplayed()) {
			Log.pass("WebElement List items: "+linkpresent.getText()+" Displayed");
			links = true;
			Weblocator.explicitWait(1);
			linkpresent.sendKeys(filePath);
		} 
		return links;



	}

	public static boolean multipleselect_and_deselect_all_selected_options(By link, String value1, String value2) throws InterruptedException{
		boolean selectoption=false;
		WebElement element = waitForOptionalElement(link);
		if (element != null && element.isDisplayed()) {
			Log.pass("WebElement List items: "+element.getText()+" Displayed");
			Select listbox = new Select(driver.findElement(link));
			//To verify that specific select box supports multiple select or not.
			if(listbox.isMultiple())
			{
				System.out.print("Multiple selections is supported");
				listbox.selectByVisibleText(value1);
				listbox.selectByValue(value2);
				listbox.selectByIndex(5);
				Thread.sleep(3000);
				//To deselect all selected options.
				listbox.deselectAll();
				Thread.sleep(2000);
			}
			selectoption=true;
		}
		return selectoption;
	}

	public static void navigate_back(){
		driver.navigate().back();
	} 

	public static void navigate_forward(){
		driver.navigate().forward();
	}
	
	public static String randomeNum(int no) {
		Log.pass("Call RandomNumber method");
		String generatedString2 = RandomStringUtils.randomNumeric(no);
		return (generatedString2);
	}
	
	public static boolean deleteFile(String filePath) {
		boolean status=false;
		File file = new File(filePath); 
        if(file.delete()) 
        { 
            System.out.println("File deleted successfully"); 
            status=true;
        } 
        else
        { 
            System.out.println("Failed to delete the file"); 
        }
		return status;
	}
	
	public static boolean getFileUpdate(String filePath, String writeData) throws IOException {
		boolean status=false;

		File FC = new File(filePath);//Created object of java File class.
		try {
			FC.createNewFile();
			FileWriter FW = new FileWriter(filePath);

			BufferedWriter BW = new BufferedWriter(FW);

			BW.write(writeData); //Writing In To File.

			BW.newLine();//To write next string on new line.

			//BW.flush();  /* ye line if you want to create CSV file vena delete or comment kr do or uppar bhi  file ka extance bhi change kr do */

			BW.close();
			status=true;

		} catch (IOException e2) {
			e2.printStackTrace();
		}

		return status;
	}


}
