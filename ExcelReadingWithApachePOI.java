package com.scratch.GuviProject;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadingWithApachePOI {

	public static void main(String[] args) throws IOException {
		
		readExcel();

	}

	public static void readExcel() throws IOException {

		FileInputStream myFile = new FileInputStream("C:\\Users\\srija\\OneDrive\\Desktop\\Maniish\\Java\\Task 08.xlsx");

		XSSFWorkbook myWorkbook = new XSSFWorkbook(myFile);

		org.apache.poi.ss.usermodel.Sheet mySheet = myWorkbook.getSheetAt(0);

		for (Row row : mySheet) {
			
			for (Cell cell : row) {
				
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue() + "\t");
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue() + "\t");
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue() + "\t");
					break;
				default:
					System.out.print("Unknown cell type\t");
					break;
				}
			}
			System.out.println();
		}
	}

}
