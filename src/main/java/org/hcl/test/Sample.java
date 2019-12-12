package org.hcl.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Sample {
	public static void main(String[] args) throws IOException {
	File loc =new File("C:\\Users\\elcot\\eclipse-vijay\\MavenSample\\test\\Worksheet.xlsx");
	FileInputStream a = new FileInputStream(loc);
	Workbook w =new XSSFWorkbook(a);
	Sheet s= w.getSheet("sheet");
	Row r = s.getRow(1);
	Cell c = r.getCell(1);
	System.out.println(c);
	
	int physicalNumberOfRows = s.getPhysicalNumberOfRows();
	System.out.println("No of Rows" +physicalNumberOfRows);

	int phys=r.getPhysicalNumberOfCells();
	System.out.println("asdfghjk"+phys);
	
	for (int i = 0; i < r .getPhysicalNumberOfCells(); i++) {	
		System.out.println(r.getCell(i));
	}
	
	

		
			}

}
