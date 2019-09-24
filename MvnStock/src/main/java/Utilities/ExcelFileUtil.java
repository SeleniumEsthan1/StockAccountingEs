package Utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.xmlbeans.impl.xb.xsdschema.Public;

public class ExcelFileUtil {
	Workbook wb;
	
	// it will load Excel Sheet
	public ExcelFileUtil() throws Throwable 
	{
		FileInputStream fis=new FileInputStream("D:\\ESTHAN\\AUTOMATION\\Automation Testing\\MvnStock\\TestInputs\\InputSheet.xlsx");
		wb=WorkbookFactory.create(fis);
		
		
	}
	//row count
	
	public int rowCount(String sheetname)
	{
		return wb.getSheet(sheetname).getLastRowNum();
	}
	//column count
	
	public int colCount(String sheetname,int row)
	{
		return wb.getSheet(sheetname).getRow(row).getLastCellNum();
	}
	//reading the data
	
	@SuppressWarnings("deprecation")
	public String getData(String sheetname, int row, int column )
	{
		String data="";
		if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC) {
		int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
		data=String.valueOf(celldata);
		
		}
		else
		{
			data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}   
	//storing data into Excel sheet pass or fail and not executed
	public void setData(String sheetname, int row, int column,String status) throws Throwable
	{
		Sheet sh=wb.getSheet(sheetname);
		Row rownum=sh.getRow(row);
		Cell cell= rownum.createCell(column);
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass"))
		{
			//create cell style
			CellStyle style=wb.createCellStyle();
			//create font
			Font font=wb.createFont();
			//apply color to text
			font.setColor(IndexedColors.GREEN.index);
			//apply bold to text
			font.setBold(true);
			//set font
			style.setFont(font);
			//set cell style
			rownum.getCell(column).setCellStyle(style);
		}
		else
		if(status.equalsIgnoreCase("Fail"))
		{
			//create cell style
			CellStyle style=wb.createCellStyle();
			//create font
			Font font=wb.createFont();
			//apply color to text
			font.setColor(IndexedColors.RED.index);
			//apply bold to text
			font.setBold(true);
			//set font
			style.setFont(font);
			//set cell style
			rownum.getCell(column).setCellStyle(style);
		} else
			
		if(status.equalsIgnoreCase("Not Executed"))
			{
				//create cell style
				CellStyle style=wb.createCellStyle();
				//create font
				Font font=wb.createFont();
				//apply color to text
				font.setColor(IndexedColors.BLUE.index);
				//apply bold to text
				font.setBold(true);
				//set font
				style.setFont(font);
				//set cell style
				rownum.getCell(column).setCellStyle(style);
			}
		FileOutputStream fos=new FileOutputStream("D:\\ESTHAN\\AUTOMATION\\Automation Testing\\MvnStock\\TestOutput\\OutputSheet.xlsx");
		wb.write(fos);
		fos.close();
		
		}
	public static void main(String args[])throws Throwable
	{
		ExcelFileUtil excel=new ExcelFileUtil();
		
		System.out.println(excel.rowCount("Sheet1"));
		System.out.println(excel.colCount("Sheet1",1));
		System.out.println(excel.getData("Sheet1",1,1));
		
		
		
		excel.setData("Sheet1", 1, 2, "Pass");
		excel.setData("Sheet1", 2, 2, "Fail");
		excel.setData("Sheet1", 3, 2, "Not Executed");
		
	}
}
