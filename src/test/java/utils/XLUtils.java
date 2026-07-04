package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLUtils 
{

	public static FileInputStream fi; 
	public static FileOutputStream fo;
	public static Workbook wb;
	public static Sheet ws;
	public static Row row;
	public static Cell cell;
	
	
	
	public static int getRowCount(String xlfile,String xlsheet) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		int rowcount = ws.getLastRowNum();
		wb.close();
		return rowcount;	
	}	
	
	public static short getColumnCount(String xlfile,String xlsheet,int rownum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		short colcount = row.getLastCellNum();
		wb.close();
		return colcount;		
	}
	
	public static String getStringData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(colnum);
		
		String data;
		try 
		{
			data = cell.getStringCellValue();
		} catch (Exception e) 
		{
			data = "";
			System.err.println("no data found!");
		}
		wb.close();
		return data;
	}
	
	public static double getNumericData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(colnum);
		
		double data;
		try 
		{
			data = cell.getNumericCellValue();
		} catch (Exception e) 
		{
			data = 0.0;
			System.err.println("no data found!");
		}
		wb.close();
		return data;
	}
	
	public static boolean getBooleanData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(colnum);
		
		boolean data;
		try 
		{
			data = cell.getBooleanCellValue();
		} catch (Exception e) 
		{
			data = false;
			System.err.println("no data found!");
		}
		wb.close();
		return data;
	}
	
	public static void setPass(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.createCell(colnum);
		cell.setCellValue("Pass");
		
		CellStyle pstyle = wb.createCellStyle();
		pstyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		pstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(pstyle);
		fo = new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
	}

	public static void setFail(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.createCell(colnum);
		cell.setCellValue("Fail");
		
		CellStyle fstyle = wb.createCellStyle();
		fstyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		fstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(fstyle);
		fo = new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
	}
}
