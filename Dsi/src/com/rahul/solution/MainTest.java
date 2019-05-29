package com.rahul.solution;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;



public class MainTest {

	public static void main(String[] args) throws InterruptedException, IOException {
		// TODO Auto-generated method stub
		String path = MainTest.class.getClassLoader().getResource("chromedriver.exe").getPath();
		System.setProperty("webdriver.chrome.driver", path);
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.internetworldstats.com/top20.htm");
		driver.manage().window().maximize();
		System.out.println(driver.getTitle());
		Thread.sleep(4000);
		int rowCount = driver.findElements(By.xpath("/html/body/table[4]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr")).size();
		System.out.println(rowCount);
		ArrayList<String> countryList = new ArrayList<String>();
		for(int i=3;i<=rowCount-4;i++) {
			String s =  driver.findElement(By.xpath("/html/body/table[4]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr["+i+"]/td[2]/p/a/font/b")).getText();
			//System.out.println(s);
			countryList.add(s);
		}
		System.out.println(countryList.size());
		
		
		WriteExcel writeExcel = new WriteExcel();
		writeExcel.initialize();
		
		searchCountry(countryList,driver);
		


		 

	}
	

	
	public static void searchCountry(List<String> countryList,WebDriver driver) throws IOException {
		
		
		int column=2;
		

        
		for(int c=0;c<countryList.size();c++) {
			
			driver.get("https://www.google.co.uk/");
			driver.findElement(By.xpath("//*[@id=\"tsf\"]/div[2]/div/div[1]/div/div[1]/input")).sendKeys(countryList.get(c));
			driver.findElement(By.xpath("//*[@id=\"tsf\"]/div[2]/div/div[3]/center/input[1]")).click();
			/*IF YOUR BROWSER DO NOT ASK FOR ENGLISH LANGUAGE SETTINGS PLEASE COMMENT THIS BELOW 2 LINES.*/
			if(c==0)
				driver.findElement(By.xpath("//*[@id=\"Rzn5id\"]/div/a[2]")).click();
			try {
				Thread.sleep(1000);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			int l = driver.findElements(By.cssSelector(".Z1hOCe")).size();
			System.out.println(l);
			CountryDetails countryDetails = new CountryDetails();
			List<Details> detailsList = new ArrayList<Details>();
		
			List<WebElement> elements = driver.findElements(By.className("Z1hOCe"));
			java.util.Iterator<WebElement> i = elements.iterator();
			while(i.hasNext()) {
			    WebElement element = i.next();
			    if (element.isDisplayed()) {
			      // Do something with the element
			    	Details details = new Details();
			    	String s = element.getText();
			    	details.setDescription(s);
					detailsList.add(details);
			    }
			} 
			countryDetails.setCountryName(countryList.get(c));
			countryDetails.setDetails(detailsList);

			
			File file = new File("C:\\Users\\HP\\Desktop\\EvaluationSheet.xlsx");
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet1 = wb.getSheetAt(0);
			
			CellStyle style = wb.createCellStyle();
	        Font font = wb.createFont();
	        font.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
	        style.setFont(font);
			
	        String s="";
			if(c<9)
				s = "Sample Data 0" + (c+1);
			else
				s = "Sample Data " + (c+1);
			System.out.println(sheet1.getDefaultRowHeight());
			sheet1.getRow(6).createCell(column).setCellValue(s);
			sheet1.getRow(8).createCell(column).setCellValue(countryDetails.getCountryName());
			sheet1.getRow(8).getCell(column).setCellStyle(style);
			int row = 9;
		
			for(Details detail: countryDetails.getDetails()) {

				sheet1.getRow(row).createCell(column).setCellValue(detail.getDescription());
				row++;
			}
			column=column+5;
			FileOutputStream out = new FileOutputStream(file); 
	        wb.write(out);
			wb.close();
		}
		 
		


	}
	
	

}
