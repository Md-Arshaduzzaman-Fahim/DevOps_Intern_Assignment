package Assignment;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class lolA {
	public static void main(String[] args) {
		WriteData("Wednesday",2,3,4,"gg","ggwp");
	}
	public static final String TEST_DATA_SHEET_PATH = "./src/test/resources/4BeatsQ1.xlsx"; 
	public static Workbook book;
	public static Sheet sheet;
	
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
