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
		int ordersCounter=0;
		//create the order object at the beginning. If the PO cell is non null, initialize the object and start filling values in it
		Order order = null;
		String uniqueOrderKey = null;
		for (int i = 0; i <= rowsCount; i++) {
			XSSFRow row = sheet.getRow(i);
			int colCounts = row.getLastCellNum();					
			for (int j = 0; j < colCounts; j++) {
				XSSFCell cell = row.getCell(j);

				// Uncomment below to print the contents on the console
				/**
				if(cell != null){

					int cellType = cell.getCellType();
					String cellValueAsString;

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

						System.out.println("CELL_TYPE_STRING[" + i + "," + j + "]=" + cell.getStringCellValue());
					}
				}
				 **/
				
				
				String cellValueAsString = null;
				int cellType = 9999;
				try {
					cellType = cell.getCellType();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}				
				
				switch (cellType) {
				case XSSFCell.CELL_TYPE_BOOLEAN:
					cellValueAsString = "" +cell.getBooleanCellValue();
					break;
				case XSSFCell.CELL_TYPE_ERROR:
					cellValueAsString = cell.getErrorCellString();
					break;
				case XSSFCell.CELL_TYPE_FORMULA:
					cellValueAsString = cell.getCellFormula();
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					cellValueAsString = "" + cell.getNumericCellValue();
					break;
				case XSSFCell.CELL_TYPE_STRING:
					cellValueAsString = cell.getStringCellValue();
					break;
				default:
					cellValueAsString = "";
					break;
				}

				//Check For PO
				if(i != 0 && j==0 && cell != null){
					order = new Order();
					ordersCounter++;					
					uniqueOrderKey = ordersCounter + props.getProperty("uniqueKeyJoiner") + cellValueAsString;
					order.setPO(cellValueAsString);
					System.out.println("Obtained an Order! With key "+uniqueOrderKey);
				}
				
				//Check For Sold To
				if(i != 0 && j==1 && cell != null){
					order.setSoldTo(cellValueAsString);
				}
				
				//Check For Ship To
				if(i != 0 && j==2 && cell != null){
					order.setShipTo(cellValueAsString);
				}
				
				//Check For Dropship Indicator
				if(i != 0 && j==3 && cell != null){
					order.setDropshipIndicator(cellValueAsString);
				}

			}
		}
		wb.close();
	}
}
