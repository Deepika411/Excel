package org.test.Sample1;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
public class Two {
	
	 // Excel file path
	 public static final String loc = "H:\\Lachu\\Testng\\Sample1\\excel\\test.xlsx";

	public static void main(String[] args) throws IOException, InvalidFormatException {
	 FileInputStream file = new FileInputStream(loc);
	 XSSFWorkbook workbook = new XSSFWorkbook(file);
	 // Getting sheet1
	 Sheet sheet = workbook.getSheet("Sheet1");
	 // Getting row at index 0 in sheet1
	 Row row = sheet.getRow(0);
	 int rowLength = row.getPhysicalNumberOfCells();
	 // Creating sheet2
	 Sheet sheetTwo = workbook.createSheet("test7");
	 // Creating row at index 0 in sheet2
	 Row sheetTwoRow = sheetTwo.createRow(0);
	 // Setting value in row of sheet2 from sheet1
	 for (int i = 0; i < rowLength; i++) {
	 Cell cell = sheetTwoRow.createCell(i);
	 Cell firstSheetCell = row.getCell(i);
	 cell.setCellValue(firstSheetCell.getStringCellValue());
	 }
	 // Writing changes in Excel file
	 file.close();
	 FileOutputStream outFile = new FileOutputStream(new File(loc));
	 workbook.write(outFile);
	 System.out.println("done");
	 outFile.close();
	 }
	 }

