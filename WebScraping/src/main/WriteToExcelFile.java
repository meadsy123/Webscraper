package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
 
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class WriteToExcelFile {
	public static final String SAMPLE_XLSX_FILE_PATH = "C:\\Users\\jakem\\OneDrive\\Documents\\prisync-links.xlsx";
	public static final String COMPETITOR_FILE_PATH = "C:\\Users\\jakem\\eclipse-workspace\\WebScraping\\files\\websites.xlsx";
 
    public static void write(List<Product> products, List<Linkname> links) throws IOException {
    	
        // create work book and sheets
        XSSFWorkbook workBook = new XSSFWorkbook();
        XSSFSheet sheet = workBook.createSheet("simple sheet");
 
        // create rows and columns
        Row rowHeader = sheet.createRow(0);
        rowHeader.createCell(0).setCellValue("Name");
        rowHeader.createCell(1).setCellValue("SKU Code");
        rowHeader.createCell(2).setCellValue("Product Cost");
        rowHeader.createCell(3).setCellValue("Brand");
        rowHeader.createCell(4).setCellValue("Category");
        rowHeader.createCell(5).setCellValue("Your Own URL");
        for(int i =0; i < links.size();i++) {
        	if((links.get(i) != null) & (links.get(i).getName() != null)){
        		rowHeader.createCell(6+i).setCellValue(links.get(i).getName());
        	}
        }
        int rowNum = 1; // header is already created
        for (Product product : products) {
            Row rowDetail = sheet.createRow(rowNum);
            rowDetail.createCell(0).setCellValue(product.getName());
            rowDetail.createCell(1).setCellValue(product.getSku());
            rowDetail.createCell(2).setCellValue(product.getCost());
            rowDetail.createCell(3).setCellValue(product.getBrand());
            rowDetail.createCell(4).setCellValue(product.getCategory());
            rowDetail.createCell(5).setCellValue(product.getNisurl());
            if(product.getComp1() != null) {
            rowDetail.createCell(6).setCellValue(product.getComp1());
            }
            if(product.getComp2() != null) {
                rowDetail.createCell(7).setCellValue(product.getComp2());
                }
            if(product.getComp3() != null) {
                rowDetail.createCell(8).setCellValue(product.getComp3());
                }
            if(product.getComp4() != null) {
                rowDetail.createCell(9).setCellValue(product.getComp4());
                }
            if(product.getComp5() != null) {
                rowDetail.createCell(10).setCellValue(product.getComp5());
                }
            if(product.getComp6() != null) {
                rowDetail.createCell(11).setCellValue(product.getComp6());
                }
            if(product.getComp7() != null) {
                rowDetail.createCell(12).setCellValue(product.getComp7());
                }
            if(product.getComp8() != null) {
                rowDetail.createCell(13).setCellValue(product.getComp8());
                }
            if(product.getComp9() != null) {
                rowDetail.createCell(14).setCellValue(product.getComp9());
                }
            if(product.getComp10() != null) {
                rowDetail.createCell(15).setCellValue(product.getComp10());
                }
            if(product.getComp11() != null) {
                rowDetail.createCell(16).setCellValue(product.getComp11());
                }
            if(product.getComp12() != null) {
                rowDetail.createCell(17).setCellValue(product.getComp12());
                }
            if(product.getComp13() != null) {
                rowDetail.createCell(18).setCellValue(product.getComp13());
                }
            if(product.getComp14() != null) {
                rowDetail.createCell(19).setCellValue(product.getComp14());
                }
            if(product.getComp15() != null) {
                rowDetail.createCell(20).setCellValue(product.getComp15());
                }
            if(product.getComp16() != null) {
                rowDetail.createCell(21).setCellValue(product.getComp16());
                }
            if(product.getComp17() != null) {
                rowDetail.createCell(22).setCellValue(product.getComp17());
                }
            ++rowNum;
        }
 
        FileOutputStream fos = null;
 
        try {
            fos = new FileOutputStream(new File(SAMPLE_XLSX_FILE_PATH));
            workBook.write(fos);
        }
 
        finally {
            fos.close();
            workBook.close();
        }
 
    }
public static void writeCompetitor(List<Competitor> competitors) throws IOException {
    	
        // create work book and sheets
        XSSFWorkbook workBook = new XSSFWorkbook();
        XSSFSheet sheet = workBook.createSheet("simple sheet");
 
        // create rows and columns
        Row rowHeader = sheet.createRow(0);
        rowHeader.createCell(0).setCellValue("Competitor Name");
        rowHeader.createCell(1).setCellValue("Url");
        rowHeader.createCell(2).setCellValue("DOM Select");
        rowHeader.createCell(3).setCellValue("DOM validation Select");
        rowHeader.createCell(4).setCellValue("Validation Type");
 
        int rowNum = 1; // header is already created
        for (Competitor competitor : competitors) {
            Row rowDetail = sheet.createRow(rowNum);
            rowDetail.createCell(0).setCellValue(competitor.getName());
            rowDetail.createCell(1).setCellValue(competitor.getUrl());
            rowDetail.createCell(2).setCellValue(competitor.getSelect());
            rowDetail.createCell(3).setCellValue(competitor.getValidation_select());
            rowDetail.createCell(4).setCellValue(competitor.getValidation_type());
            ++rowNum;
        }
 
        FileOutputStream fos = null;
 
        try {
            fos = new FileOutputStream(new File(COMPETITOR_FILE_PATH));
            workBook.write(fos);
        }
 
        finally {
            fos.close();
            workBook.close();
        }
 
    }
public static void writeProduct (Product product) throws IOException {
	
	//read excel spreadsheet
		FileInputStream excel = new FileInputStream(new File(SAMPLE_XLSX_FILE_PATH));
		
	// create work book and sheets
    XSSFWorkbook workbook = new XSSFWorkbook(excel);
	//Get first sheet from the workbook
  	XSSFSheet sheet = workbook.getSheetAt(0);
  	
  	// get current row and append data
    int rowLast =  sheet.getLastRowNum();

    int rowNum = rowLast + 1; // add a new row
    
        Row rowDetail = sheet.createRow(rowNum);
        rowDetail.createCell(0).setCellValue(product.getName());
        rowDetail.createCell(1).setCellValue(product.getSku());
        rowDetail.createCell(2).setCellValue(product.getCost());
        rowDetail.createCell(3).setCellValue(product.getBrand());
        rowDetail.createCell(4).setCellValue(product.getCategory());
        rowDetail.createCell(5).setCellValue(product.getNisurl());
        if(product.getComp1() != null) {
            rowDetail.createCell(6).setCellValue(product.getComp1());
            }
            if(product.getComp2() != null) {
                rowDetail.createCell(7).setCellValue(product.getComp2());
                }
            if(product.getComp3() != null) {
                rowDetail.createCell(8).setCellValue(product.getComp3());
                }
            if(product.getComp4() != null) {
                rowDetail.createCell(9).setCellValue(product.getComp4());
                }
            if(product.getComp5() != null) {
                rowDetail.createCell(10).setCellValue(product.getComp5());
                }
            if(product.getComp6() != null) {
                rowDetail.createCell(11).setCellValue(product.getComp6());
                }
            if(product.getComp7() != null) {
                rowDetail.createCell(12).setCellValue(product.getComp7());
                }
            if(product.getComp8() != null) {
                rowDetail.createCell(13).setCellValue(product.getComp8());
                }
            if(product.getComp9() != null) {
                rowDetail.createCell(14).setCellValue(product.getComp9());
                }
            if(product.getComp10() != null) {
                rowDetail.createCell(15).setCellValue(product.getComp10());
                }
            if(product.getComp11() != null) {
                rowDetail.createCell(16).setCellValue(product.getComp11());
                }
            if(product.getComp12() != null) {
                rowDetail.createCell(17).setCellValue(product.getComp12());
                }
            if(product.getComp13() != null) {
                rowDetail.createCell(18).setCellValue(product.getComp13());
                }
            if(product.getComp14() != null) {
                rowDetail.createCell(19).setCellValue(product.getComp14());
                }
            if(product.getComp15() != null) {
                rowDetail.createCell(20).setCellValue(product.getComp15());
                }
            if(product.getComp16() != null) {
                rowDetail.createCell(21).setCellValue(product.getComp16());
                }
            if(product.getComp17() != null) {
                rowDetail.createCell(22).setCellValue(product.getComp17());
                }
    
    excel.close();

    FileOutputStream fos = null;

    try {
        fos = new FileOutputStream(new File(SAMPLE_XLSX_FILE_PATH));
        workbook.write(fos);
    }

    finally {
        fos.close();
        workbook.close();
    }

}
public static void writeHeader (List<Linkname> linknames) throws IOException {

	//read excel spreadsheet
	FileInputStream excel = new FileInputStream(new File(SAMPLE_XLSX_FILE_PATH));
	
	//Get the workbook instance for XLS file 
	XSSFWorkbook workbook = new XSSFWorkbook(excel);
	
	//Get first sheet from the workbook
	XSSFSheet sheet = workbook.getSheetAt(0);
		
	// create rows and columns
	Row rowHeader = sheet.createRow(0);
	
	//add data to the cells
	rowHeader.createCell(0).setCellValue("Name");
	rowHeader.createCell(1).setCellValue("SKU Code");
	rowHeader.createCell(2).setCellValue("Product Cost");
	rowHeader.createCell(3).setCellValue("Brand");
	rowHeader.createCell(4).setCellValue("Category");
	rowHeader.createCell(5).setCellValue("Your Own URL");
	for(int i =0; i < linknames.size();i++) {
    	if((linknames.get(i) != null) & (linknames.get(i).getName() != null)){
    		rowHeader.createCell(6+i).setCellValue(linknames.get(i).getName());
    	}
    }
	excel.close();
	FileOutputStream fos = null;

    try {
        fos = new FileOutputStream(new File(SAMPLE_XLSX_FILE_PATH));
        workbook.write(fos);
    }

    finally {
    	fos.close();
        workbook.close();
    }
 
} 

}