package com.qa.excel.operations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelIterator {

	public static void main(String[] args) throws Exception {
		String excelFilePath = "E:\\eclipse-workspace\\Apache_POI\\DataFiles\\Countries.xlsx";

		FileInputStream fp = new FileInputStream(excelFilePath);
	try (XSSFWorkbook workbook = new XSSFWorkbook(fp)) {
		//XSSFSheet sheet = workbook.getSheet(Sheet1);	
		 XSSFSheet sheet = workbook.getSheetAt(0);
	
		 int rows = sheet.getLastRowNum();
		 System.out.println("Last row: "+rows);
		 
		 int cols = sheet.getRow(0).getLastCellNum();
		 System.out.println("No. of columns: "+cols);
	
	// Iterator
	
	Iterator iterator = sheet.iterator();
	while(iterator.hasNext())
	{
		XSSFRow row =(XSSFRow) iterator.next();
		
		Iterator cellIterator = row.cellIterator();
  while(cellIterator.hasNext()) {
	  
	  XSSFCell cell = (XSSFCell) cellIterator.next();
  
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
	
}}
}