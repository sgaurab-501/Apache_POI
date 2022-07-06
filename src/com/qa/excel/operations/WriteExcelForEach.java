package com.qa.excel.operations;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelForEach {

	public static void main(String[] args) throws Exception {
		
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet sheet =  workbook.createSheet("Emp Info");
			Object empData[][] = {{"EmpId", "Name", "Job"}, 
					{"101", "Ram", "Engineer"}, 
			    	{"102", "Mohan", "Manager"},               
			    	{"103", "Aditya", "Analyst"},	
			    	{"104", "Sumit", "Doctor"},
			};
			
			//Using For each loop
			
			int rowCount = 0;
			
			for(Object emp[]: empData) {
				
			XSSFRow row = sheet.createRow(rowCount++);
			
			int colCount=0;
			 
			for(Object value: emp) {
				XSSFCell cell = row.createCell(colCount++);	
			
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
		}	
		
System.out.println("employee.xlsx file created successfully........");

}
	
	
	
}
