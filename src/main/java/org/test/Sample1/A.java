package org.test.Sample1;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class A {
	public static WebDriver driver;
	
	 public void main(String args[]) throws Exception
	   {
		 System.setProperty("webdriver.chrome.driver", "C:\\Users\\HP\\Downloads\\chromedriver_win32\\chromedriver.exe");
			WebDriver wb = new ChromeDriver();    
		    A tp =new A();
	        tp.setup();
	        tp.Handle_Dynamic_Webtable();
	        tp.tearDown();

	    }
	    public void setup() throws Exception 
	    {
	        driver.manage().window().maximize();
	        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
	        driver.get("http://www.moneycontrol.com/");
	    }

	  
	    public void tearDown() throws Exception {
	        driver.quit();
	    }

	    
	
	    public void Handle_Dynamic_Webtable() throws FileNotFoundException 
	    {
	        FileOutputStream fos = new FileOutputStream("H:\\Lachu\\Testng\\Sample1\\excel\\writeto.xlsx");                                 
	        XSSFWorkbook wkb = new XSSFWorkbook();       
	        XSSFSheet sheet1 = wkb.createSheet("DataStorage");
	        WebElement mytable = driver.findElement(By.xpath(".//*[@id='mc_mainWrapper']/section/div/div[2]/aside/div[3]/div[2]/div[1]/table"));
	        List<WebElement> rows_table = mytable.findElements(By.tagName("tr"));
	        int rows_count = rows_table.size();
	        System.out.println("Number of Rows " + rows_count);
	        for (int row = 0; row < rows_count; row++) 
	        { 
	            XSSFRow excelRow = sheet1.createRow(row); 
	            if(row==0)
	            {
	                List<WebElement> head_row = rows_table.get(row).findElements(By.tagName("th"));
	                int Head_count = head_row .size();
	                System.out.println("Number of Header cells In Row 0 are "+ Head_count);
	                for(int i=0;i<Head_count;i++) 
	                {
	                        XSSFCell excelCell = excelRow.createCell(i);
	                        excelCell.setCellType(XSSFCell.CELL_TYPE_STRING);
	                        String celtext = head_row.get(i).getText();
	                        excelCell.setCellValue(celtext);
	                        System.out.println("Header in valuein column number " + i + " Is " + celtext);
	                }
	                
	            }
	            else
	            {
	                List<WebElement> Columns_row = rows_table.get(row).findElements(By.tagName("td"));
	                int columns_count = Columns_row.size();
	                System.out.println("Number of cells In Row " + row + " are "+ columns_count);
	                for (int column = 0; column < columns_count; column++) 
	                {
	                    XSSFCell excelCell = excelRow.createCell(column);
	                    excelCell.setCellType(XSSFCell.CELL_TYPE_STRING);
	                    String celtext = Columns_row.get(column).getText();
	                    excelCell.setCellValue(celtext);
	                    System.out.println("Cell Value Of row number " + row+ " and column number " + column + " Is " + celtext);
	                }
	    
	            }
	            System.out.println("-----------------------------------");
	        }
	    
	        try {
	            fos.flush();
	            wkb.write(fos);
	            fos.close();
	        }
	        catch(Exception e)
	        {
	        	e.printStackTrace();
	        }
	    }
}


	    
