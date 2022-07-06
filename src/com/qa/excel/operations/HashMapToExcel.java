package com.qa.excel.operations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.record.aggregates.DataValidityTable;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMapToExcel {

	public static void main(String[] args) throws IOException {
	
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet =  workbook.createSheet("Student Data");

	Map<String, String> data = new HashMap<String, String>();
	
	data.put("101", "Ram");
	data.put("102", "Mohan");	
	data.put("103", "Ajit");
	data.put("104", "Ramesh");
	data.put("105", "Rohit");
	
	int rowNum = 0;
	
for(Map.Entry<String, String> entry:data.entrySet()) {
	
	XSSFRow row = sheet.createRow(rowNum++);
	
	row.createCell(0).setCellValue((String)entry.getKey());
	row.createCell(1).setCellValue((String)entry.getValue());
	
	}

	FileOutputStream fout = new FileOutputStream(".\\DataFiles.\\Student.xlsx");
	workbook.write(fout);
	fout.close();
	
	System.out.println("Excel file written successfully..........");
	
	}
}