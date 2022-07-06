package com.qa.excel.operations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelData {

	public static void main(String[] args) throws Exception {
	
	
			XSSFWorkbook workbook = new XSSFWorkbook();	

			
		
	XSSFSheet sheet =  workbook.createSheet("Emp Info");
	Object empData[][] = {{"EmpId", "Name", "Job"}, 
			{"101", "Ram", "Engineer"}, 
        	{"102", "Mohan", "Manager"},               
        	{"103", "Aditya", "Analyst"},	
        	{"104", "Sumit", "Doctor"},
	};

// Using For Loop
	
int rows = empData.length;	
int cols = empData[0].length;	
	
System.out.println("Rows: "+rows);
System.out.println("Columns: "+cols);

for(int r=0; r<rows;r++) {
	
	XSSFRow row = sheet.createRow(r);
	
	for(int c=0; c<cols;c++) {
		
     XSSFCell cell = row.createCell(c);
	Object value = empData[r][c];
	
if(value instanceof String) 	
	cell.setCellValue((String)value);

if(value instanceof Integer) 	
	cell.setCellValue((Integer)value);

if(value instanceof Boolean) 
	cell.setCellValue((Boolean)value);
	}}

String excelFilePath = 	".\\DataFiles\\employee.xlsx";
FileOutputStream fout = new FileOutputStream(excelFilePath);
workbook.write(fout);
	
fout.close();
	System.out.println("employee.xlsx file created successfully........");

	}
	
	
}
