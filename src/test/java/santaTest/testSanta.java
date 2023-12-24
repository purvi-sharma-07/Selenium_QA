package santaTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class testSanta {

	public static void main(String[] args) {
		
	     

        // Initialize ChromeDriver
		
        WebDriver driver = new ChromeDriver();

        // Navigate to the webpage
        driver.get("file:///C:/Users/user/Documents/santa.html");

        // Locate the UI element and get the text
        driver.findElement(By.id("name")).sendKeys("Ps");
        driver.findElement(By.id("SS")).sendKeys("Harshita");
        WebElement element = driver.findElement(By.xpath("//button[text()='Submit']"));
        element.click();
        String textFromUILine1 = driver.findElement(By.tagName("p")).getText().split("\n")[0];
        String textFromUILine2 = driver.findElement(By.tagName("p")).getText().split("\n")[1];
        //WebElement element1 = driver.findElement(By.id("output"));
        //String textFromUI = element1.getText();

        // Write the text to an Excel sheet
        writeToExcel(textFromUILine1,textFromUILine2);

        // Close the browser
        driver.quit();
    }

    private static void writeToExcel(String text, String text1) {
        // Create a new workbook
    	
    	File file = new File("E:\\output1.xlsx");
    	XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a sheet 
        XSSFSheet sheet = workbook.createSheet("sheet1");

        // Create a row
        sheet.createRow(1).createCell(0).setCellValue(text);
        sheet.createRow(2).createCell(0).setCellValue(text1);
        //Row row = sheet.createRow(0);

        // Create a cell and set the text
        //Cell cell = row.createCell(0);
        //cell.setCellValue(text);

        try {
            // Write the workbook to an Excel file
            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            System.out.println("Text written to Excel successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}


