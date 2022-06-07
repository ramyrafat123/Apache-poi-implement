package excelOperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) {

		String excelFilePath=".\\DataFile\\testdata.xlsx";
		//		stream data from file ,we use class of fileinputStream

		try {
			FileInputStream inputstream=new FileInputStream(excelFilePath);


			//			find workbook 
			try {
				XSSFWorkbook workbook=new XSSFWorkbook(inputstream);

				//				find sheet from workbook
				XSSFSheet sheet=workbook.getSheetAt(0);

				//				Using for loop to get rows/cells from sheet
				int rows=sheet.getLastRowNum();
				int cols=sheet.getRow(1).getLastCellNum();

//				for (int i = 0; i < rows; i++) {
//
//					//					get data of each row and put it in row object
//					XSSFRow row=sheet.getRow(i);
//
//					for (int j = 0; j < cols; j++) {
//						//					get data of each cell and put it in col object	
//						XSSFCell col=row.getCell(j);
//						switch (col.getCellType()) {
//						case STRING: System.out.print(col.getStringCellValue());
//						break;
//						case NUMERIC: System.out.print(col.getNumericCellValue());
//						break;
//						case BOOLEAN: System.out.print(col.getBooleanCellValue());
//						break;
//
//						}
//						
//					}
//					System.out.println();
//
//				}
				
//				using iterator to read data from excel sheet
				
			Iterator iterator=sheet.iterator();
			
			while (iterator.hasNext()) {
				XSSFRow row = (XSSFRow) iterator.next();
				Iterator celliterator=row.cellIterator();
				
				while (celliterator.hasNext()) {
					XSSFCell cell = (XSSFCell) celliterator.next();
					switch (cell.getCellType()) {
					case STRING: System.out.print(cell.getStringCellValue());
					break;
					case NUMERIC: System.out.print(cell.getNumericCellValue());
					break;
					case BOOLEAN: System.out.print(cell.getBooleanCellValue());
					break;

					}
					
					
					
				}
				System.out.println();
				
				
			}

			} catch (IOException e) {


				e.printStackTrace();
			}

		} catch (FileNotFoundException e) {

			e.printStackTrace();
		}

	}

}
