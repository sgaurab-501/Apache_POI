package com.qa.excel.operations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DateCellOperation {

	public static void main(String[] args) throws IOException {
		
	XSSFWorkbook workbook = new XSSFWorkbook();
	
	XSSFSheet sheet = workbook.createSheet("Date Formats");
	
    XSSFRow row = sheet.createRow(0);
    XSSFCell cell = row.createCell(0);
    
    // Date in number format
    
    cell.setCellValue(new Date());
    
    
    CreationHelper crtHelper = workbook.getCreationHelper();
    
  // Format 1: dd-mm-yyyy  
    
    CellStyle style1 = workbook.createCellStyle();
     style1.setDataFormat(crtHelper.createDataFormat().getFormat("dd-mm-yyyy"));
    
     XSSFCell cell1 = sheet.createRow(1).createCell(0);
     cell1.setCellValue(new Date());
     cell1.setCellStyle(style1);
    
 // Format 2: mm-dd-yy
    
    CellStyle style2 = workbook.createCellStyle();
    style2.setDataFormat(crtHelper.createDataFormat().getFormat("mm-dd-yy"));
    XSSFCell cell2 = sheet.createRow(2).createCell(0);
    cell2.setCellValue(new Date());
    cell2.setCellStyle(style2);
    
// Format 3: mm-dd-yyyy hh:mm:ss    
    
    CellStyle style3 = workbook.createCellStyle();
    style2.setDataFormat(crtHelper.createDataFormat().getFormat("mm-dd-yyyy hh:mm:ss")); // Specify the date format
    XSSFCell cell3 = sheet.createRow(3).createCell(0);
    cell3.setCellValue(new Date());
    cell3.setCellStyle(style3);
    
    
// Format 4: hh:mm:ss    
    
    CellStyle style4 = workbook.createCellStyle();
    style2.setDataFormat(crtHelper.createDataFormat().getFormat("hh:mm:ss"));
    XSSFCell cell4 = sheet.createRow(4).createCell(0);
    cell4.setCellValue(new Date());
    cell4.setCellStyle(style4);
    
       
    String excelFilePath = ".\\DataFiles\\DateFormat.xlsx";
    
    FileOutputStream fout = new FileOutputStream(excelFilePath);
	
	workbook.write(fout);
	workbook.close();
	fout.close();
	
	System.out.println("File created successfully.............");
	
	
	
	
	}
	
	
	
}
