package working_with_excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelStuffWithTwoDimensionalArray {


	
	public static void main(String[] args) throws Exception {
		
		 String[][] result = getAllSheetData("MOCK_DATA.xlsx", "data");
		 System.out.println(Arrays.deepToString(result));
		 
		System.out.println( getCellIntRowIndex("MOCK_DATA.xlsx", "data", 1, 2));
	}
	
	public static String[][] getAllSheetData(String fileName, String sheetName) throws Exception {

		FileInputStream fis = new FileInputStream(fileName);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet(sheetName);
		int rowCount = sh.getPhysicalNumberOfRows();
		int colCount = sh.getRow(0).getLastCellNum();
		//									[11]      [6]	
		String[][] dataHolder = new String[rowCount][colCount];
		
		for (int i = 0; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {
				Cell c = sh.getRow(i).getCell(j);
				dataHolder[i][j] = c.toString();		
			}
		}
		wb.close();
		fis.close();
		return dataHolder;
	}
	
	public static String getCellIntRowIndex(String fileName, String sheetName, int rowIndex, int colIndex ) throws Exception {
	
//		FileInputStream fis = new FileInputStream(fileName);
//		Workbook wb = WorkbookFactory.create(fis);
//		Sheet sh = wb.getSheet(sheetName);
//		int rowCount = sh.getPhysicalNumberOfRows();
//		int colCount = sh.getRow(0).getLastCellNum();
//		String[][] dataHolder = new String[rowCount][colCount];
//		
//		for (int i = 0; i < rowCount; i++) {
//			for (int j = 0; j < colCount; j++) {
//				Cell c = sh.getRow(i).getCell(j);
//				dataHolder[i][j] = c.toString();
//			}
//		}
//		
//		if(rowIndex < 12 && colIndex < 7) {
//			return dataHolder[rowIndex][colIndex];
//		}else {
//			return "Does not exist";
//		}
		//OR ALTERNATIVELY AND EASILY
		String[][] data = getAllSheetData(fileName, sheetName);
		return data[rowIndex][colIndex];
	}

}
