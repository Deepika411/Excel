package org.test.Sample1;
import java.util.*; 
import  java.io.*;  
import org.apache.poi.xssf.usermodel.XSSFCell; 
import org.apache.poi.xssf.usermodel.XSSFRow; 
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
import org.openqa.selenium.By; 
import org.openqa.selenium.WebDriver; 
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;  
public class WebTable {
	public static void main(String[] args) throws IOException  
	{     
	System.out.println("Hello Dear.....");     
	System.out.println();  
System.setProperty("webdriver.chrome.driver", "C:\\Users\\HP\\Downloads\\chromedriver_win32\\chromedriver.exe");
	WebDriver wb = new ChromeDriver();          
	wb.navigate().to("http://www.w3schools.com/html/html_tables.asp"); 
	wb.manage().window().maximize();     
	System.out.println(wb.getTitle() +" - WebPage has been launched");   

	List<WebElement> irows =   wb.findElements(By.xpath("//*[@id='main']/table[1]/tbody/tr"));     
	int iRowsCount = irows.size();     
	List<WebElement> icols =   wb.findElements(By.xpath("//*[@id='main']/table[1]/tbody/tr[1]/th"));     
	int iColsCount = icols.size();     
	System.out.println("Selected web table has " +iRowsCount+ " Rows and " +iColsCount+ " Columns");     
	System.out.println();      

	FileOutputStream fos = new FileOutputStream("H:\\Lachu\\Testng\\Sample1\\excel\\writeto.xlsx");                                 

	XSSFWorkbook wkb = new XSSFWorkbook();       
	XSSFSheet sheet1 = wkb.createSheet("DataStorage"); 

	for (int i=1;i<=iRowsCount;i++)      
	{               
	for (int j=1; j<=iColsCount;j++)                    
	{           
	if (i==1)       
	{           
	WebElement val= wb.findElement(By.xpath("//*[@id='main']/table[1]/tbody/tr["+i+"]/th["+j+"]"));             
	String  a = val.getText();            
	System.out.print(a);                        

	XSSFRow excelRow = sheet1.createRow(i); 
	
	XSSFCell excelCell = excelRow.createCell(j);                  
	excelCell.setCellType(XSSFCell.CELL_TYPE_STRING);                 
	excelCell.setCellValue(a);  

	//wkb.write(fos);       
	}       
	else        
	{           
	WebElement val= wb.findElement(By.xpath("//*[@id='main']/table[1]/tbody/tr["+i+"]/td["+j+"]"));             
	String a = val.getText();                    
	System.out.print(a);                            

	XSSFRow excelRow = sheet1.createRow(i);             
	XSSFCell excelCell = excelRow.createCell(j);                      
	excelCell.setCellType(XSSFCell.CELL_TYPE_STRING);                   
	excelCell.setCellValue(a);   

	//wkb.write(fos);       
	}       
	}               
	System.out.println();     
	}     
	fos.flush();     
	wkb.write(fos);     
	fos.close();     
	}
	}

