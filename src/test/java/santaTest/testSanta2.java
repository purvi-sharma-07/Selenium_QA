package santaTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class testSanta2 {

	public static void main(String[] args) {
	     
	        // Initialize ChromeDriver
			
	        WebDriver driver = new ChromeDriver();

	        // Navigate to the webpage
	        driver.get("file:///C:/Users/user/Documents/button.html");

	        // Locate the UI element and get the text
	        String[] names= {"PS","AK","BM","TS","HK","jl","MM"};
	        
	        WebElement allnames=driver.findElement(By.id("name"));
	        File file = new File("E:\\output1.xlsx");
	    	XSSFWorkbook workbook = new XSSFWorkbook();
	        XSSFSheet sheet = workbook.createSheet("sheet1");
	        Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Secret Santa");
	        for(int i=0;i< names.length;i++) {
	        	
	        allnames.sendKeys(names[i]);
	        WebElement element = driver.findElement(By.xpath("//button[text()='Submit']"));
	 	    element.click();
	 	    driver.findElement(By.xpath("//button[text()='Choose your Secret Santa']")).click(); 
	        String e1=driver.findElement(By.xpath("//div[@id='output']")).getText();
	        String e2=driver.findElement(By.xpath("//div[@id='output1']")).getText();
	        Row row = sheet.createRow(i + 1);
	    	 row.createCell(0).setCellValue(e1);
		     row.createCell(1).setCellValue(e2);
	        allnames.clear();
	        }
	        driver.quit();
	     
	        try {
	            // Write the workbook to an Excel file
	            FileOutputStream outputStream = new FileOutputStream(file);
	            workbook.write(outputStream);
	            workbook.close();
	            outputStream.close();
	            System.out.println("Text in Excel is written successfully.");
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }

}



