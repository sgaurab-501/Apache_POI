package com.qa.excel.operations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateFormulaCell {

	public static void main(String[] args) throws IOException {
		
	String	excelFilePath = ".\\DataFiles.\\BookPrice.xlsx";
		FileInputStream fin = new FileInputStream(excelFilePath);
		
	XSSFWorkbook workbook = new XSSFWorkbook(fin);
		
	XSSFSheet sheet = workbook.getSheetAt(0);
//	the count of rows start from zero. Hence, the total no. of rows will always come as actual-1

	int rows = sheet.getLastRowNum();
	System.out.println("Total no. of rows: "+rows);
	
// the getCell() formula wil count the cell starting from zero
	
	sheet.getRow(5).getCell(2).setCellFormula("SUM(C2:C5)");	
	
	fin.close();
	
	FileOutputStream fout= new FileOutputStream(excelFilePath);
	
	workbook.write(fout);
	
	fout.close();
	System.out.println("Formula updated in sheet successfully...........");
	
	
	
	}
	
	
	
}
