package com.qa.excel.operations;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws Exception {
		
//	String excelFilePath = "E:\\eclipse-workspace\\Apache_POI\\DataFiles\\Countries.xlsx";

	String excelFilePath = ".\\DataFiles\\Countries.xlsx";
	
	
		FileInputStream fp = new FileInputStream(excelFilePath);
	 
		XSSFWorkbook workbook = new XSSFWorkbook(fp);
		
		//XSSFSheet sheet = workbook.getSheet(Sheet1);	
		 XSSFSheet sheet = workbook.getSheetAt(0);
		 
		 
		 //**********Using For loop***********************
		 
		 int rows = sheet.getLastRowNum();
		 System.out.println("Last row: "+rows);
		 
		 int cols = sheet.getRow(0).getLastCellNum();
		
		 
		 System.out.println("No. of columns: "+cols);
		
		 for(int r=0; r<rows; r++) {
			
	     XSSFRow  row =sheet.getRow(r);
		  for(int c=0; c<cols; c++) {
			
		XSSFCell cell =	row.getCell(c);
		
		 switch(cell.getCellType())
		{
		case STRING: System.out.print(cell.getStringCellValue());
			break;
			
		case NUMERIC: System.out.print(cell.getNumericCellValue());
			break;
			
		case BOOLEAN: System.out.print(cell.getBooleanCellValue());	
			break;
		
		default: System.out.print("Invalid data in cell");
		break;
		}
		 
		 System.out.print(" | ");
			}
			
			System.out.println();
			
		}
	}

}