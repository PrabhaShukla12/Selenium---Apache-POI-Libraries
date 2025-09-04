package apachePOI;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ReadExcel {

	public static void main(String[] args) throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream("C:\\Users\\HP\\Downloads\\Selenium\\Data.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis); 
		
		XSSFSheet sh= wb.getSheet("Data");
		
		int rowcount = sh.getLastRowNum();
		
		System.out.println("Total Rows in the Excel file are: "+ rowcount);
		
		
		System.setProperty("webdriver.chrome.driver","C:\\Users\\HP\\Downloads\\chromedriver-win64 (2)\\chromedriver-win64\\chromedriver.exe");
		WebDriver d=new ChromeDriver();
		
		d.manage().window().maximize();
		
		for(int i=0; i<= rowcount; i++)
		{
		d.get("https://www.facebook.com/");	
		
		d.findElement(By.id("email")).sendKeys(sh.getRow(i).getCell(0).getStringCellValue());
		
		d.findElement(By.id("pass")).sendKeys(sh.getRow(i).getCell(1).getStringCellValue());
		Thread.sleep(1000);
			
		System.out.print(sh.getRow(i).getCell(0).getStringCellValue());
		
		System.out.println(" "+ sh.getRow(i).getCell(1).getStringCellValue());
		
		//System.out.println(sh.getRow(i).getCell(2).getNumericCellValue()); 
		}
		
		d.findElement(By.name("login")).click();
		
	}

}
