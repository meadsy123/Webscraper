/**
 * 
 */
package main;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
/**
 * @author jakem
 *
 */

public class ReadExcelFile {
    public static final String SAMPLE_XLSX_FILE_PATH = "C:\\Users\\jakem\\OneDrive\\Documents\\product.xlsx";
    public static final String COMPETITOR_FILE_PATH = "C:\\Users\\jakem\\eclipse-workspace\\WebScraping\\files\\competitors.xlsx";

    public List<Product> read() throws IOException, InvalidFormatException {
    	
    	//read excel spreadsheet
    	FileInputStream file = new FileInputStream(new File(SAMPLE_XLSX_FILE_PATH));
		
		//Get the workbook instance for XLS file 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		//Get first sheet from the workbook
      	XSSFSheet sheet = workbook.getSheetAt(0);

        List<Product> products = new ArrayList<Product>();
        // for-each loop to iterate over the rows and columns
        
        for (Row row: sheet) {
        	//make sure row is not the header row
        	if(row.getRowNum() > 0) {
        		Product product = new Product();
        		if (row.getCell(0) != null) {
        			product.setName(row.getCell(0).toString());
        		}
        		if (row.getCell(1) != null) {
        			product.setSku(row.getCell(1).toString());
        		}
        		if (row.getCell(2) != null) {
        			product.setCost(row.getCell(2).toString());
        		}
        		if (row.getCell(3) != null) {
        			product.setBrand(row.getCell(3).toString());
        		}
        		if (row.getCell(4) != null) {
        			product.setCategory(row.getCell(4).toString());
        		}
        		if (row.getCell(5) != null) {
        			product.setNisurl(row.getCell(5).toString());
        		}
        		
        		products.add(product);
        	} 	
	    }
        file.close();
        workbook.close();
        return products;
    }
    
public List<Competitor> getCompetitors() throws IOException, InvalidFormatException {
    	
    	//read excel spreadsheet
    	FileInputStream file = new FileInputStream(new File(COMPETITOR_FILE_PATH));
		
		//Get the workbook instance for XLS file 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		//Get first sheet from the workbook
      	XSSFSheet sheet = workbook.getSheetAt(0);

        List<Competitor> competitors = new ArrayList<Competitor>();
        // for-each loop to iterate over the rows and columns
        
        for (Row row: sheet) {
        	//make sure row is not the header row
        	if(row.getRowNum() > 0) {
        		Competitor competitor = new Competitor();
        		if (row.getCell(0) != null) {
        			competitor.setName(row.getCell(0).toString());
        		}
        		if (row.getCell(1) != null) {
        			competitor.setUrl(row.getCell(1).toString());
        		}
        		if (row.getCell(2) != null) {
        			competitor.setSelect(row.getCell(2).toString());
        		}
        		if (row.getCell(3) != null) {
        			competitor.setValidation_select(row.getCell(3).toString());
        		}
        		if (row.getCell(4) != null) {
        			competitor.setValidation_type(row.getCell(4).toString());
        		}
        		if (row.getCell(5) != null) {
        			competitor.setConnection_type(row.getCell(5).toString());
        		}
        		
        		competitors.add(competitor);
        	} 	
	    }
        file.close();
        workbook.close();
        return competitors;
    }
}
