package com.learn;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderMain {

	public static void main(String[] args) throws IOException, InvalidFormatException {

		ExcelReaderMain exlRdr = new ExcelReaderMain();
		ExcelTasksMain exlTasks = new ExcelTasksMain();
		double sumTotal = 0;
		boolean status = false;
		
		ClassLoader loader = exlRdr.getClass().getClassLoader();
		File file = new File(loader.getResource("input.xlsx").getFile());
		System.out.println(file.getAbsolutePath());
		
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		sumTotal = exlTasks.getTotal(workbook);
		System.out.println("Sum total : "+sumTotal);
		workbook.close();
		
		//status = exlTasks.writeToExcel(sumTotal);
		
		if(status)
			System.out.println("Task completed successfully!");
		else
			System.out.println("Writing not done!!");
	}
	
}
