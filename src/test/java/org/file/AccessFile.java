package org.file;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.text.DateFormatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AccessFile {
	
	public static void main(String[] args) throws IOException {
	
	
	File f = new File("E:\\Eclipse\\DataDriven2\\excel\\NEWDATA.xlsx");
	
	FileInputStream fi = new FileInputStream(f);
	
	Workbook w = new XSSFWorkbook(fi);
	
	Sheet sheet = w.getSheet("Login");
	
	Row row = sheet.getRow(1);
	
	Cell cell = row.getCell(3);
	
	int cellType = cell.getCellType();
	
	if(cellType == 1) {
		String value = cell.getStringCellValue();
		System.out.println(value);
	}
	else {
		if(DateUtil.isCellDateFormatted(cell)) {
			Date dateCellValue = cell.getDateCellValue();
			SimpleDateFormat s = new SimpleDateFormat("yyyy-MMM-dd");
			String value = s.format(dateCellValue);
			System.out.println(value);
		}
		else {
			double numericCellValue = cell.getNumericCellValue();
			long l = (long)numericCellValue;
			String value = String.valueOf(l);
			System.out.println(value);
			
		}
	}
	
	
	}
	
}
