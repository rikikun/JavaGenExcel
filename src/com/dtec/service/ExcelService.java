package com.dtec.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelService {

	public void writeExcel(XSSFWorkbook workBook,String fileName) throws IOException{
		FileOutputStream outPut = new FileOutputStream(new File(fileName));
		workBook.write(outPut);
		outPut.close();
	}
	
	public XSSFWorkbook readExcel(String fileName) throws IOException{
		FileInputStream file = new FileInputStream(new File(fileName));
		return new XSSFWorkbook(file);
	}
}
