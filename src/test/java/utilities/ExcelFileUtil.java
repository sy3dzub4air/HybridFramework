package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil{
	Workbook wb;
	//constructor for reading excel file
	public ExcelFileUtil(String ExcelPath) throws Throwable
	{
		FileInputStream fi = new FileInputStream(ExcelPath);
		wb = WorkbookFactory.create(fi);
	}
	// Count no of Rows in a sheet
	public int rowCount(String SheetName)
	{
		return wb.getSheet(SheetName).getLastRowNum();
	}
	// Method for reading Cell Data
	public String getCellData(String SheetName, int row, int column)
	{
		String data;
		if(wb.getSheet(SheetName).getRow(row).getCell(column).getCellType()== CellType.NUMERIC)
		{
			int celldata = (int) wb.getSheet(SheetName).getRow(row).getCell(column).getNumericCellValue();
			data = String.valueOf(celldata);
		}
		else
		{
			data = wb.getSheet(SheetName).getRow(row).getCell(column).getStringCellValue();
		}
		return data;

	}
	// method for set cell data
	public void setCellData(String SheetName, int row, int columns, String status, String writeExelpath) throws Throwable
	{
		// get sheet from workbook wb
		Sheet ws = wb.getSheet(SheetName);
		// get row from  sheet
		Row rowNum = ws.getRow(row);
		// create cell in a row
		Cell cell = rowNum.createCell(columns);
		//Write status
		cell.setCellValue(status);
		if (status.equalsIgnoreCase("Pass"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			// color with green
			font.setColor(IndexedColors.GREEN.getIndex());;
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(columns).setCellStyle(style);

		}
		else if (status.equalsIgnoreCase("fail"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			// color with red
			font.setColor(IndexedColors.RED.getIndex());;
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(columns).setCellStyle(style);

		}
		else if (status.equalsIgnoreCase("Blocked"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			// color with Blue
			font.setColor(IndexedColors.BLUE.getIndex());;
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(columns).setCellStyle(style);

		}

		FileOutputStream fo = new FileOutputStream(writeExelpath);
		wb.write(fo);
	}
	public static void main(String[] args) throws Throwable
	{
		// create object for file
	ExcelFileUtil xl = new ExcelFileUtil("Z:/Sample Data.xlsx");
	int rc = xl.rowCount("Login Data");
	System.out.println(rc);
	
	// read all rows 
	for (int i = 1; i <=rc; i++)
	{
	String fname = xl.getCellData("Login Data", i, 0);
	String mname = xl.getCellData("Login Data", i, 1);
	String lname = xl.getCellData("Login Data", i, 2);
	String eid = xl.getCellData("Login Data", i, 3);
	System.out.println(fname+"   "+mname+"   "+lname+"   "+eid);
	// write as pass into cell
	//xl.setCellData("Login Data", i, 4, "PASS", "Z:/Reports1.xlsx");
	//xl.setCellData("Login Data", i, 4, "FAIL", "Z:/Reports1.xlsx");
	xl.setCellData("Login Data", i, 4, "BLOCKED", "Z:/Reports1.xlsx");
	}
	
	}
}
