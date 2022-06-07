package excelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// create workbook-->Sheet-->row-->cell
public class WritingExcel {

	public static void main(String[] args) throws IOException {
	
//		create workbook
		
		XSSFWorkbook workbook=new XSSFWorkbook();
//		create sheet inside workbook
		
		XSSFSheet sheet=workbook.createSheet("Emp Inf");
		
//		prepare data for add in excelsheet
		
		Object empdata[][]= {
				{"EmpId","Name","Job"},
				{"1","ramy","QA"},
				{"2","ahmed","QA"},
				{"3","mahmoud","QA"}
		};
		
		int rows=empdata.length;
		int cells=empdata[0].length;
		
		for (int i = 0; i < rows; i++) {
		XSSFRow	row=sheet.createRow(i);
			
			for (int j = 0; j < cells; j++) {
				
				XSSFCell cell=row.createCell(j);
				
				Object value =empdata[i][j];
				
				if(value instanceof String)
					cell.setCellValue((String) value);
				if(value instanceof Integer)
					cell.setCellValue((Integer) value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean) value);
			}
			
		}
		
		String filePath=".//Datafile//employee.xlsx";
		FileOutputStream outputStream=new FileOutputStream(filePath);
		
		workbook.write(outputStream);
		
		outputStream.close();
		
		System.out.println("success");
		

	}

}
