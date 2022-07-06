package com.qa.excel.operations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelArrayList {

	public static void main(String[] args) throws IOException {
		
XSSFWorkbook workbook = new XSSFWorkbook();
XSSFSheet sheet =  workbook.createSheet("Emp Info");	
		

ArrayList<Object[]> empData = new ArrayList<Object[]>();
empData.add(new Object[]{"EmpId", "Name", "Job"});
empData.add(new Object[]{"101", "Ram", "Engineer"});
empData.add(new Object[]{"102", "Mohan", "Manager"});
empData.add(new Object[]{"103", "Aditya", "Analyst"});
empData.add(new Object[]{"104", "Sumit", "Doctor"});

// Using For each loop

int rowNum=0;
for(Object[] emp: empData) {
XSSFRow row =	sheet.createRow(rowNum++);
 
int cellNum=0;
	for(Object value: emp) {
		XSSFCell cell = row.createCell(cellNum++);	
		
		if(value instanceof String) 	
			cell.setCellValue((String)value);
	
		if(value instanceof Integer) 	
			cell.setCellValue((Integer)value);
	
		if(value instanceof Boolean) 
			cell.setCellValue((Boolean)value);

	}
	}
	
String excelFilePath = 	".\\DataFiles\\employee.xlsx";
FileOutputStream fout = new FileOutputStream(excelFilePath);
workbook.write(fout);

fout.close();
System.out.println("employee.xlsx file created successfully........");	
}

}
