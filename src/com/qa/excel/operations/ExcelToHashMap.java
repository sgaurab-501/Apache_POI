package com.qa.excel.operations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToHashMap {
	
public static void main(String[] args) throws IOException {
	
	String excelPath = ".\\DataFiles.\\Student.xlsx";
	FileInputStream fin = new FileInputStream(excelPath);
	
	XSSFWorkbook workbook = new XSSFWorkbook(fin);
	
	XSSFSheet sheet = workbook.getSheet("Student Data");
	
	int rows = sheet.getLastRowNum();
	System.out.println("Total rows: "+rows);
	
	
	HashMap<String, String> data = new HashMap<String, String>();
	
//  Reading data from excel to HashMap
	
	for(int r=0; r<=rows; r++) {
		
 String key = sheet.getRow(r).getCell(0).getStringCellValue();
 String value = sheet.getRow(r).getCell(1).getStringCellValue();

 data.put(key, value);
 
 	}
	
// Read Data from HashMap
	
	
	for(Map.Entry<String, String> entry: data.entrySet()) {
		
		
		System.out.println(entry.getKey()+" | "+entry.getValue());
		
	}
	
	
	
}

}
