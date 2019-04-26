package org.test.Sample1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.Normalizer.Form;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Duplicates 
{
	public static void main(String[] args) throws IOException  
	{ 
	 try 
	 {           
	     FileInputStream file = new FileInputStream(new File("H:\\Lachu\\Testng\\Sample1\\excel\\writeto.xlsx"));  
	     List sheetData = new ArrayList();

	    XSSFWorkbook workbook = new XSSFWorkbook(file); 

	    XSSFSheet sheet = workbook.getSheetAt(0);
	  ArrayList<Form> vipList = new ArrayList<Form>();
	    Iterator<Row> rowIterator = sheet.iterator();   
	    while (rowIterator.hasNext()) 
	    {            
	        Row row = rowIterator.next();

	        Iterator<Cell> cellIterator = row.cellIterator();   
	        List data = new ArrayList();

	        while (cellIterator.hasNext())  
	        { 

	            Cell cell = cellIterator.next();    

	            switch (cell.getCellType())                     
	            {        
	                case Cell.CELL_TYPE_NUMERIC:  System.out.print(cell.getNumericCellValue() + "\t"); 
	            break;                       
	                case Cell.CELL_TYPE_STRING: System.out.print(cell.getStringCellValue() + "\t");  
	            break;     
	            }           
	        }

	    }  
	 }
	 finally
	 {
		 System.out.println("done");
	 }
	
	}
}



	    

