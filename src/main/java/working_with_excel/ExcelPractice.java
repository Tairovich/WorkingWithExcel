package working_with_excel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelPractice {

	public static void main(String[] args) throws Exception {

		// Workbook -> Sheet -> Row -> Cell
		/*
		 * Earlier versions of poi library have 2 different sets of classes to work with
		 * xls, xlsx file
		 * 
		 * xls files aka MS Excel 1997-2003 were handles by these clases below
		 * HSSFWorkbook, HSSFSheet, HSSFRow, HSSFCell
		 * 
		 * xlsx => XSSFWorkbook, XSSFSheet, XSSFRow, XSSFCell classes
		 * 
		 */

		File excelFile = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(excelFile);
		System.out.println(wb.getNumberOfSheets());

		Sheet sh = wb.getSheetAt(0);
		System.out.println("Number of rows: " + sh.getLastRowNum());

		Row row1 = sh.getRow(0);

		Cell c1 = row1.getCell(1);

		System.out.println(c1);

		System.out.println("Number of columns: " + row1.getLastCellNum());

		// getPhysicalRpwNumber will return actual number of rows
		// whether you have empty value row or not
		int getNonEmptryRowCount = sh.getPhysicalNumberOfRows();
		System.out.println(getNonEmptryRowCount);

		int row = sh.getPhysicalNumberOfRows();
		int col = row1.getLastCellNum();

//		for (int i = 0; i < row; i++) {
//			System.out.println("ROW NUMBER: " + (i + 1));
//			Row r = sh.getRow(i);
//			for (int j = 0; j < col; j++) {
//				Cell c = r.getCell(j);
//				System.out.print(c + " ");
//			}
//			System.out.println();
//		}

		wb.close();
		
		
		
		
		
		System.out.println(getAllSheetData("MOCK_DATA.xlsx", ""));
		

	}

	public static String[][] getAllSheetData(String fileName, String sheetName) throws Exception {

		String[][] dataHolder = new String[11][6];

		File file = new File(fileName);
		Workbook wb = WorkbookFactory.create(file);
		Sheet sh = wb.getSheetAt(0);

		for (int i = 0; i < sh.getPhysicalNumberOfRows(); i++) {
			Row r = sh.getRow(i);
			 
			for (int j = 0; j < r.getLastCellNum(); j++) {
				Cell c = r.getCell(j);
				dataHolder[i][j] = c.toString();
				System.out.print(dataHolder[i][j]  + "  ");		
			}
			System.out.println();
		}
		wb.close();
		return dataHolder;
	}

}
