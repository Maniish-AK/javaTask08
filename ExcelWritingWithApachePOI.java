package com.scratch.GuviProject;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWritingWithApachePOI {

	public static void main(String[] args) {
		
		writeExcel();

	}

	public static void writeExcel() {

		XSSFWorkbook myWorkbook = new XSSFWorkbook();

		org.apache.poi.ss.usermodel.Sheet mySheet = myWorkbook.createSheet("Sheet1");
		
		Row row = mySheet.createRow(0);
		Cell cell1 = row.createCell(0);
		cell1.setCellValue("Name");
		Cell cell2 = row.createCell(1);
		cell2.setCellValue("Age");
		Cell cell3 = row.createCell(2);
		cell3.setCellValue("Email");
		
		Row row1 = mySheet.createRow(1);
		row1.createCell(0).setCellValue("James Bond");
		row1.createCell(1).setCellValue("40");;
		row1.createCell(2).setCellValue("james@test.com");
		
		Row row2 = mySheet.createRow(2);
		row2.createCell(0).setCellValue("John Wick");
		row2.createCell(1).setCellValue("35");;
		row2.createCell(2).setCellValue("john@test.com");
		
		Row row3 = mySheet.createRow(3);
		row3.createCell(0).setCellValue("Tom Cruise");
		row3.createCell(1).setCellValue("30");;
		row3.createCell(2).setCellValue("tom@test.com");
		
		mySheet.autoSizeColumn(0);
		mySheet.autoSizeColumn(1);
		mySheet.autoSizeColumn(2);
		
		try(FileOutputStream fileOut = new FileOutputStream("Task 08 Write.xlsx")) {
			
			myWorkbook.write(fileOut);
		} catch(IOException e) {
			
			e.printStackTrace();
		}
		
		System.out.println("Excel file created successfully!");

	}
}
