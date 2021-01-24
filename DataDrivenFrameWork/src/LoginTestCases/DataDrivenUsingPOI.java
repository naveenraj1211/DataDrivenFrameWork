package LoginTestCases;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class DataDrivenUsingPOI 
{
	@Test
	public void readExcel() throws FileNotFoundException,IOException
	{
		FileInputStream excel = new FileInputStream("C:\\Users\\naveenraj.durairaj\\Downloads\\testleaf.xlsx");
		Workbook workbook = new XSSFWorkbook(excel);
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row>  rowIterator = sheet.rowIterator();
		
		while(rowIterator.hasNext()) 
		{
			Row rowValue=rowIterator.next();
			Iterator<Cell> columnIterator=rowValue.iterator();
			
			while(columnIterator.hasNext())
			{
				Cell cellValue=columnIterator.next();
				cellValue.getStringCellValue();
				System.out.println(cellValue);
			}			
		}
		
		

		
	}
	
	
	
	
}
