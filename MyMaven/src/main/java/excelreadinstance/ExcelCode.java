package excelreadinstance;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode {
	FileInputStream file;
	XSSFWorkbook work;
	XSSFSheet sheet;
	public ExcelCode() throws IOException
	{
		file=new FileInputStream("C:\\Users\\JITHIN NAIR\\Desktop\\priya.xlsx");
		work=new XSSFWorkbook(file);
		sheet=work.getSheet("data");
	}
	
	public String readStringData(int row ,int column)
	{
		
		Row r=sheet.getRow(row);
		Cell c=r.getCell(column);
		int  CellType =c.getCellType();
		switch(CellType)
		{
		case Cell.CELL_TYPE_STRING:
		{
			return c.getStringCellValue();
			
			}
		case Cell.CELL_TYPE_NUMERIC:
		{
			int a=(int) c.getNumericCellValue();
			return String.valueOf(a);
		}
		}
		return null;
	
		
	}
	
	

}
