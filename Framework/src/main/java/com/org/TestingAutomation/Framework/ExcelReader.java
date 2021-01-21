package com.org.TestingAutomation.Framework;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public FileInputStream  fis = null;
	 String path = null;
	 private XSSFWorkbook workbook = null;
	 private XSSFSheet sheet =null;
	 private XSSFRow row = null;
	 private XSSFCell cell = null;
	 public FileOutputStream fileout =null;
	 
	 public ExcelReader() throws IOException {
		 path = System.getProperty("D:\\SeleniumWebdriver-Java-from-Scratch\\Framework\\testdata\\testdat.xlsx");
		 fis = new FileInputStream(path);
		 workbook = new XSSFWorkbook(fis);
		 sheet = workbook.getSheetAt(0);
		 		 
	 }
	 
	 //Total number of rows in a sheet
	 public int getSheetRow(String SheetName) {
		int index= workbook.getSheetIndex(SheetName);
		sheet = workbook.getSheetAt(index);
		return (sheet.getLastRowNum()+1);
				 
	 }
	 
	 
	//Total number of column in a sheet
	 public int getSheetColumns(String SheetName) {
		 int index= workbook.getSheetIndex(SheetName);
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			return(row.getLastCellNum());
			 
			
	 }
	 
	 public String getCellData(String SheetName,int ColNum, int rownum) {
		 int index= workbook.getSheetIndex(SheetName);
			sheet = workbook.getSheetAt(index);
			
			row = sheet.getRow(rownum);
			cell =row.getCell(ColNum);
			
			return (cell.getStringCellValue());		 
	 }
	 
	 public String getCellData(String SheetName,String ColNum, int rownum) {
		 int Colnum=-1;
		 int index= workbook.getSheetIndex(SheetName);
			sheet = workbook.getSheetAt(index);
			
			for(int i=0;i<=getSheetColumns(SheetName);i++) {
				row = sheet.getRow(0);
				cell = row.getCell(i);
				
				if(cell.getStringCellValue().equals(ColNum)) {
					Colnum= cell.getColumnIndex();
					break;
				}
				 
			}
			row =sheet.getRow(rownum);
			cell =row.getCell(Colnum);
			return (cell.getStringCellValue());
	 }
	 
	 //
	 public void setCellData(String sheetName,int colNum,int rowNum,String str) {
		 int index = workbook.getSheetIndex(sheetName);
		 sheet = workbook.getSheetAt(index);
		 row = sheet.getRow(rowNum);
		 cell =row.createCell(colNum);
		 cell.setCellValue(str);
		 try {
			fileout = new FileOutputStream(path);
			try {
				workbook.write(fileout);
				fileout.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 
	 }
	 
	 public static void main(String args[]) throws IOException {
		 ExcelReader reader = new ExcelReader();
		 reader.getCellData("SignUp", 0, 2);
	 }
}

