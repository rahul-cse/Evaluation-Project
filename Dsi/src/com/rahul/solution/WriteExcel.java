package com.rahul.solution;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {
	
	
	public void initialize() throws IOException {
		File file = new File("C:\\Users\\HP\\Desktop\\EvaluationSheet.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 = wb.getSheetAt(0);
		sheet1.createRow(6);
		sheet1.createRow(8);
		sheet1.createRow(9);
		sheet1.createRow(10);
		sheet1.createRow(11);
		sheet1.createRow(12);
		sheet1.createRow(13);
		sheet1.createRow(14);
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		wb.close();
	}
	
	
	


	
	
	
}
