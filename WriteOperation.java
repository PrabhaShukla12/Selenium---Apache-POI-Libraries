package apachePOI;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteOperation {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		FileInputStream fis = new FileInputStream("C:\\Users\\HP\\Downloads\\Selenium\\Data.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis); 
		
		XSSFSheet sh= wb.getSheet("Data");
		
		XSSFRow r0= sh.createRow(0);
		
		XSSFCell c0 = r0.createCell(0);
		c0.setCellValue("Name");

		XSSFCell c1 = r0.createCell(1);
		c1.setCellValue("Class");
		
		XSSFCell c2 = r0.createCell(2);
		c2.setCellValue("Marks");
		
		XSSFRow r1= sh.createRow(1);
		
		XSSFCell c11 = r1.createCell(0);
		c11.setCellValue("Bharath");

		XSSFCell c12 = r1.createCell(1);
		c12.setCellValue("Class V");
		
		XSSFCell c13 = r1.createCell(2);
		c13.setCellValue("567");
		
		FileOutputStream fos = new FileOutputStream("C:\\Users\\HP\\Downloads\\Selenium\\Data.xlsx");
		
		wb.write(fos);
		System.out.println("End of Writing");
	}	

}
