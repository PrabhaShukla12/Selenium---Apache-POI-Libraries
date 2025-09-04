package apachePOI;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class ReadExcellNew {

	public static void main(String[] args) throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		FileInputStream fis = new FileInputStream("C:\\Users\\HP\\Downloads\\Selenium\\Data.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis); 
		
		XSSFSheet sh= wb.getSheet("Data");
		
		int rowcount = sh.getLastRowNum();
		
		System.out.println("Total Rows in the Excel file are: "+ rowcount);
		
	}
	
	}
