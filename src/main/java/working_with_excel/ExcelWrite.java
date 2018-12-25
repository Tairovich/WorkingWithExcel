package working_with_excel;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelWrite {

	
	public static void main(String[] args) throws Exception {
		
		
		File file = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(file);
		Sheet sh = wb.getSheetAt(0);
		Row row = sh.getRow(1);
		Cell c1 = row.getCell(1);
		
		c1.setCellValue("My OWN VALUE");
		
		FileOutputStream fos = new FileOutputStream("myown.xlsx");
		wb.write(fos);
		
		fos.close();
		wb.close();
		
		
		
		
	}
}
