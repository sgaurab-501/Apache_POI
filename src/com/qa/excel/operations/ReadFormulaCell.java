package com.qa.excel.operations;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFormulaCell {
	
	
	public static void main(String[] args) throws Exception {
		
		String excelFilePath = ".\\DataFiles\\Salary.xlsx";
		FileInputStream fin = new FileInputStream(excelFilePath);
	
	XSSFWorkbook workbook = new XSSFWorkbook(fin);
	
	XSSFSheet sheet = workbook.getSheetAt(0);
	
	int rows = sheet.getLastRowNum();
	
	int cols = sheet.getRow(0).getLastCellNum();
	
	for(int r=0; r<rows; r++) {
		
		XSSFRow row =  sheet.getRow(r);
		
		
	for(int c=0; c<cols; c++) {
		
		XSSFCell cell = row.getCell(c);
		 
	switch(cell.getCellType())
			{
			case STRING: System.out.print(cell.getStringCellValue());
				break;
				
			case NUMERIC: System.out.print(cell.getNumericCellValue());
				break;
				
			case BOOLEAN: System.out.print(cell.getBooleanCellValue());	
				break;
	
			case FORMULA: System.out.println(cell.getNumericCellValue());			
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

