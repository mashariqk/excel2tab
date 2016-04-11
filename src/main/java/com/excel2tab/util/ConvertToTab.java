package com.excel2tab.util;

import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Properties;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertToTab {

	public static void main (String[] args) throws Exception {

		Properties props = new Properties();
		File excelFile =null;
		ArrayList<Order> orders;
		try{
			InputStream resourceStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("project.properties");
			props.load(resourceStream);
		}catch(Exception e){
			e.printStackTrace();
		}

		try {
			excelFile = new File(props.getProperty("inputExcel"));
		} catch (Exception e) {
			e.printStackTrace();
		}
		XSSFWorkbook wb = new XSSFWorkbook(excelFile);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowsCount = sheet.getLastRowNum();
		System.out.println("rowsCount is "+rowsCount);
		for (int i = 0; i <= rowsCount; i++) {
			XSSFRow row = sheet.getRow(i);
			int colCounts = row.getLastCellNum();
			System.out.println("Total Number of Cols: " + colCounts);
			for (int j = 0; j < colCounts; j++) {
				XSSFCell cell = row.getCell(j);
				if(cell != null){

					int cellType = cell.getCellType();
					String cellValueAsString;

					//Print the contents on the console
					if(XSSFCell.CELL_TYPE_BLANK == cellType){
						cellValueAsString = null;
						System.out.println("CELL_TYPE_BLANK[" + i + "," + j + "]= ''");
					} else if(XSSFCell.CELL_TYPE_BOOLEAN == cellType){
						cellValueAsString = "" +cell.getBooleanCellValue();
						System.out.println("CELL_TYPE_BOOLEAN[" + i + "," + j + "]=" + cell.getBooleanCellValue());
					} else if(XSSFCell.CELL_TYPE_ERROR == cellType){
						cellValueAsString = cell.getErrorCellString();
						System.out.println("CELL_TYPE_ERROR[" + i + "," + j + "]=" + cell.getErrorCellString());
					} else if(XSSFCell.CELL_TYPE_FORMULA == cellType){
						cellValueAsString = cell.getCellFormula();
						System.out.println("CELL_TYPE_FORMULA[" + i + "," + j + "]=" + cell.getCellFormula());
					} else if(XSSFCell.CELL_TYPE_NUMERIC == cellType){
						if(DateUtil.isCellDateFormatted(cell)){
							SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
							cellValueAsString = sdf.format(cell.getDateCellValue());
						}else{
							cellValueAsString = "" + cell.getNumericCellValue();
						}                		
						System.out.println("CELL_TYPE_NUMERIC[" + i + "," + j + "]=" + cell.getNumericCellValue());
					} else{
						if(DateUtil.isCellDateFormatted(cell)){
							SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
							cellValueAsString = sdf.format(cell.getDateCellValue());
						}else{
							cellValueAsString = "" + cell.getStringCellValue();
						}
						System.out.println("CELL_TYPE_STRING[" + i + "," + j + "]=" + cell.getStringCellValue());
					}
				}				
			}
		}
	}
}
