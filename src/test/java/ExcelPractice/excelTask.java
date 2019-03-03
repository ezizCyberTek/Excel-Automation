package ExcelPractice;

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class excelTask {

	public static void main(String[] args) throws Exception {
		
		String[][] result = excelTask.getAllSheetData("MOCK_DATA.xlsx", "data");
		System.out.println(Arrays.deepToString(result));
		
		System.out.println(getCellData("MOCK_DATA.xlsx", "data", 3, 2));

	}
	
	//Create a utility method to store all sheetData in 2 dimensional String Array and return the array
		public static String[][] getAllSheetData(String filePath, String sheetName) throws Exception {
			
			//File excelFile = new File(fileName);
			FileInputStream fis = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fis);
			Sheet sh = wb.getSheet(sheetName);
			int nonEmptyRowCount = sh.getPhysicalNumberOfRows();
			int columnCount = sh.getRow(0).getLastCellNum();
			
			String[][] newArray = new String[nonEmptyRowCount][columnCount];
			
			for(int i=0; i<nonEmptyRowCount; i++) {
				
				for (int j = 0; j < columnCount; j++) {
					
					Cell cell = sh.getRow(i).getCell(j);
					newArray[i][j] = cell.toString();
					//System.out.print(cell.toString());
				}
				//System.out.println();
			}
			
			fis.close();
			wb.close();
			return newArray;
			
		}
		
		public static String getCellData(String filePath, String sheetName, int rowIndex, int colIndex) throws Exception {
			
			//one way
//			FileInputStream fis = new FileInputStream(filePath);
//			Workbook wb = WorkbookFactory.create(fis);
//			Sheet sh = wb.getSheet(sheetName);
//			
//			String dataToReturn = sh.getRow(rowIndex).getCell(colIndex).toString();
//			
//			fis.close();
//			wb.close();
//			return dataToReturn;
			
			//second way, by reusing 'getAllSheetData' method
			String[][] result = excelTask.getAllSheetData(filePath, sheetName);
			return result[rowIndex][colIndex];
			
		}

}
