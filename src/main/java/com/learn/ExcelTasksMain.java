package com.learn;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTasksMain {

	public double getTotal(XSSFWorkbook workbook){
		
		double total = 0;
		
		XSSFSheet worksheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = worksheet.rowIterator();
		
		while(rowIterator.hasNext()){
			XSSFRow row = (XSSFRow)rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			
			while(cellIterator.hasNext()){
				XSSFCell cell = (XSSFCell)cellIterator.next();
				System.out.println(cell.toString());
				switch (cell.getCellType()){
					case Cell.CELL_TYPE_NUMERIC:
						double var = cell.getNumericCellValue();
						System.out.println("Numeric type.."+var);
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.println("String type..");
						break;
					default:
						System.out.println("default type..");
				}
			}
			
		}
		
		return total; 
	}

	public boolean writeToExcel(double value){
		
		boolean status = false;
		
		return status;
	}
}
