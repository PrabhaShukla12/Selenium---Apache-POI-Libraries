import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewExcelWrite {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		FileInputStream fis= new FileInputStream("C:\\Users\\HP\\Downloads\\Selenium\\Data.xlsx");
		
		XSSFWorkbook wb= new XSSFWorkbook(fis);
		
		XSSFSheet sh= wb.getSheet("Sheet3");
		
		XSSFRow r0= sh.createRow(0);
		
		XSSFCell c0= r0.createCell(0);
		c0.setCellValue("Employee Name");
		
		XSSFCell c1= r0.createCell(1);
		c1.setCellValue("Salary");
		
		XSSFCell c2= r0.createCell(2);
		c2.setCellValue("City");
		
		XSSFRow r1= sh.createRow(1);
		
		XSSFCell c11= r1.createCell(0);
		c11.setCellValue("Ritika");
		
		XSSFCell c12= r1.createCell(1);
		c12.setCellValue("30,000 per month");
		
		XSSFCell c13= r1.createCell(2);
		c13.setCellValue("Delhi");
		
		XSSFRow r2 = sh.createRow(2);
		
		XSSFCell C21 =r2.createCell(0);
		C21.setCellValue("Arun");
	
		XSSFCell C22 =r2.createCell(1);
		C22.setCellValue("49,000 per month");
		
		XSSFCell C23 =r2.createCell(2);
		C23.setCellValue("Mumbai");
		
		XSSFRow r3 = sh.createRow(3);
		
		XSSFCell C31 =r3.createCell(0);
		C31.setCellValue("Rajat");
	
		XSSFCell C32 =r3.createCell(1);
		C32.setCellValue("75,000 per month");
		
		XSSFCell C33 =r3.createCell(2);
		C33.setCellValue("Gurgaon");
		
		XSSFRow r4 = sh.createRow(4);
		
		XSSFCell C41 =r4.createCell(0);
		C41.setCellValue("Prabha");
	
		XSSFCell C42 =r4.createCell(1);
		C42.setCellValue("2.5 lacs per month");
		
		XSSFCell C43 =r4.createCell(2);
		C43.setCellValue("Gurgaon");
		
		FileOutputStream fos= new FileOutputStream("C:\\Users\\HP\\Downloads\\Selenium\\Data.xlsx");
		wb.write(fos);
		
		System.out.println("End of Writing");
	
	}

}
