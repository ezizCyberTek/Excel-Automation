package ExcelPractice;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

public class WorkingExcel {

	public static void main(String[] args) throws Exception {
		
		File excelFile = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(excelFile);
		System.out.println(wb.getNumberOfSheets()); // gets number of sheets 
		
		Sheet sh = wb.getSheet("data"); //getting sheet with sheet name
		//Sheet sh = wb.getSheetAt(1); // getting sheet with sheet number
		
		Row row1 = sh.getRow(0); // getting row with index 0
		Cell cell1 = row1.getCell(0); // getting column with index 0
		System.out.println(cell1);
		
		int columnCount = row1.getLastCellNum(); // Column count
		System.out.println("columnCount: " + columnCount);
		
		int rowCount = sh.getLastRowNum(); // row count 
		System.out.println("row Count including empty rows : " + rowCount);
		
		//this one used most
		int nonEmptyRowCount = sh.getPhysicalNumberOfRows(); // row count (exclude empty rows)	
		System.out.println("row Count excluding empty rows : " + nonEmptyRowCount);
		
		//looping through excel sheet
		for(int i=0; i<nonEmptyRowCount; i++) {
			System.out.println("ROW NUMBER: " + (i+1));
			Row row = sh.getRow(i);
			
			for (int j = 0; j < columnCount; j++) {
				Cell cell = row.getCell(j);
				System.out.print(cell + " ");
			}
			System.out.println();
		}
		
		wb.close();

	}
	
}
