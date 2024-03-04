package com.datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {
public static void main(String[] args) throws InvalidFormatException, IOException {
	
	//excel read
	
//	String path = ("D:\Mathan\Test_For_Anni\DataDriven\src\test\resources\cc.xlsx");
	File file = new File("D:\\Mathan\\Test_For_Anni\\DataDriven\\src\\test\\resources\\cc.xlsx");
	FileInputStream inputStream = new FileInputStream(file);
	Workbook workbook = new  XSSFWorkbook(inputStream);	
//	Sheet sheet = workbook.getSheet("Sheet1");
//	Row row = sheet.getRow(1);
//	Cell cell = row.getCell(0);
//	System.out.println(cell);
//	
	//excel write
	
	Sheet createSheet = workbook.createSheet("rrrr");
	Row createRow = createSheet.createRow(0);
	Cell createCell = createRow.createCell(0);
	createCell.setCellValue("riyakutty");
	
	FileOutputStream outputStream = new FileOutputStream(file);
	workbook.write(outputStream);
	
	
}
	
}
