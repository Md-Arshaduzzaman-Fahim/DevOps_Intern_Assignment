package Assignment;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.util.List;



public class AssignmentQ1 {

	public static void main(String[] args) throws InterruptedException {
		DayOfWeek dayOfWeek = DayOfWeek.from(LocalDate.now());
		String s = dayOfWeek.toString().toLowerCase();
		String s1 = s.substring(0, 1).toUpperCase();
		String s2 = s.substring(1);
		String sheet_name = s1+s2;
		System.out.println(sheet_name);
		Object data[][] = getData(sheet_name);
		ArrayList <String> arr = new ArrayList<String>();
		for(int i=1; i<data.length; i++) {
			for(int j=0; j< data[0].length; j++) {
				if( j==2 && !data[i][j].equals(null)) {
				System.out.println(data[i][j]);
				arr.add((String) data[i][j]);
				}
			}
		}
		
		for(int i=0; i<arr.size()-1;i++) {
		
		WebDriver driver;
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");

		driver = new ChromeDriver(options);
		driver.get("https://www.google.com");
		driver.findElement(By.linkText("English")).click();
		driver.findElement(By.className("gLFyf")).sendKeys(arr.get(i));
		Thread.sleep(200);
		List <WebElement> search_result = driver.findElements(By.xpath("//div[@class='eIPGRd']//div[@class='wM6W7d']/span"));
		ArrayList <String> sr = new ArrayList <String>();
		for(WebElement e:search_result) {
			sr.add(e.getText());
		}
		sr.remove(sr.size()-1);
		System.out.println(sr);
		for (int i1=1 ;i1<sr.size(); i1++)
	    {
	        String temp = sr.get(i1);
	 
	        
	        int j = i1 - 1;
	        while (j >= 0 && temp.length() < sr.get(j).length())
	        {
	            sr.set(j+1, sr.get(j));
	            j--;
	        }
	        sr.set(j+1, temp);
	    }
        
		WriteData(sheet_name, i+2,3,4,sr.get(sr.size()-1), sr.get(0));
		
		driver.close();
	  }
	}
	
	public static final String TEST_DATA_SHEET_PATH = "./src/test/resources/4BeatsQ1.xlsx"; 
	public static Workbook book;
	public static Sheet sheet;
	
	public static Object[][] getData(String sheetName) {
		
		Object data[][] = null;
		
		try {
			FileInputStream ip = new FileInputStream(TEST_DATA_SHEET_PATH);
			book = WorkbookFactory.create(ip);
			sheet = book.getSheet(sheetName);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		for(int i=0; i<sheet.getLastRowNum(); i++) {
			for(int j=0; j<sheet.getRow(0).getLastCellNum(); j++) {
				data[i][j] = sheet.getRow(i+1).getCell(j).toString();
			}
		}
		
		return data;


}
	
	
	
	public static void WriteData(String sheetName, int rowNum, int celNum1, int celNum2, String input1, String input2) {
		
		
		
		try {
			FileInputStream ip = new FileInputStream(TEST_DATA_SHEET_PATH);
			book = WorkbookFactory.create(ip);
			sheet = book.getSheet(sheetName);
			Row r = sheet.getRow(rowNum);
			Cell c1 = r.createCell(celNum1);
			c1.setCellValue(input1);
			Cell c2 = r.createCell(celNum2);
			c2.setCellValue(input2);
			
			FileOutputStream op = new FileOutputStream(TEST_DATA_SHEET_PATH);
			book.write(op);
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
				


}

	
	}
