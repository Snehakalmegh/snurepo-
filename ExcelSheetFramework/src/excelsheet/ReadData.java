package excelsheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Cookie;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;


public class ReadData {

		
		public WebDriver driver;
		
		
		  public static void main(String[] args) throws IOException {
		  
		  //LOad Xlsx File
		  
		  FileInputStream fis=new FileInputStream("C:\\Users\\lahol\\OneDrive\\Desktop\\TB.xlsx");
		  
		  //Load Work Book 
		  XSSFWorkbook wb=new XSSFWorkbook(fis);
		  
		  XSSFSheet sheet=wb.getSheet("Sheet3");
		  
		  XSSFRow row=sheet.getRow(6);
		  
		  XSSFCell col=row.getCell(0);
		 String value=col.getStringCellValue();
		  
		 System.out.println(value);
	 
		  XSSFRow row1=sheet.getRow(6);
		  
		  XSSFCell col1=row1.getCell(2);
		  
		  String value1=col1.getStringCellValue();
		  
		  System.out.println(value1);
		  int rowcount=sheet.getLastRowNum(); System.out.println(rowcount);
		  
		  int totalrows=rowcount+1; System.out.println("Total no of Rows"+ totalrows);
		   int colcount=sheet.getRow(rowcount).getLastCellNum();
		  System.out.println(colcount);
		  for (int i = 0; i < totalrows; i++) {
		  
		  for (int j = 0; j < colcount; j++) {
		  
		  System.out.println(sheet.getRow(i).getCell(j)); } }
		  
		  
		  
		  }
		 

		@BeforeSuite
		public void openBrowser() {
			System.out.println("under Before Suite");

			System.setProperty("webdriver.chrome.driver",
					"E:\\Users\\lahol\\Downloads\\chromedriver.exe");

			driver = new ChromeDriver();
			System.out.println("Suuccess fully open Browser");

		}

		@BeforeTest
		public void enterUrl() {
			System.out.println("Under Before Test");
			
			driver.get("http://newtours.demoaut.com/");

		}

		@BeforeClass
		public void maximizewindow() {
			System.out.println("Under Before Class");
			driver.manage().window().maximize();
		}

		@BeforeMethod
		public void getcookies() {
			System.out.println("Under Before Method");
			Set<Cookie> cookies = driver.manage().getCookies();
			for (Cookie cookie : cookies) {

				System.out.println(cookie.getName());
			}
		}

		
	@Test
	public void loginTest() throws IOException
	{

		FileInputStream fis=new FileInputStream("C:\\Users\\lahol\\OneDrive\\Desktop\\TB.xlsx");
		
		//Load Work Book
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		
		XSSFSheet sheet=wb.getSheet("Batch2");
		
		XSSFSheet sheet1=wb.getSheet("Batch1");
		XSSFRow row=sheet.getRow(1);
		
		XSSFCell col=row.getCell(0);
		
		String value=col.getStringCellValue();
		
		System.out.println(value);
		
		
	   XSSFRow row1=sheet.getRow(1);
		
		XSSFCell col1=row1.getCell(1);
		
		String value1=col1.getStringCellValue();
		
		System.out.println(value1);

		System.out.println("Login With Valid user Under Test annotaion");

		driver.findElement(By.xpath("//*[@name=\"userName\"]")).sendKeys(value);
		driver.findElement(By.xpath("//*[@name=\"password\"]")).sendKeys(value1);
		driver.findElement(By.xpath("//*[@name=\"login\"]")).click();

		System.out.println("Successful Login");
	 
	}
		
}
