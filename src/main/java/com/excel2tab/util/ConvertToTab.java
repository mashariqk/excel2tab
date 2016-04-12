package com.excel2tab.util;

import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertToTab {



	public static void main (String[] args) throws Exception {

		Properties props = new Properties();
		File excelFile =null;
		ArrayList<Order> orders = null;
		try{
			InputStream resourceStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("project.properties");
			props.load(resourceStream);
		}catch(Exception e){
			e.printStackTrace();
		}

		int POColumnIndex = Integer.parseInt(props.getProperty("POColumnIndex"));
		int soldToColumnIndex = Integer.parseInt(props.getProperty("soldToColumnIndex"));
		int shipToColumnIndex = Integer.parseInt(props.getProperty("shipToColumnIndex"));
		int dropshipIndicatorColumnIndex = Integer.parseInt(props.getProperty("dropshipIndicatorColumnIndex"));
		int dropshipPoColumnIndex = Integer.parseInt(props.getProperty("dropshipPoColumnIndex"));
		int requestedDeliveryColumnIndex = Integer.parseInt(props.getProperty("requestedDeliveryColumnIndex"));
		int internalNotesColumnIndex = Integer.parseInt(props.getProperty("internalNotesColumnIndex"));
		int productCodeColumnIndex = Integer.parseInt(props.getProperty("productCodeColumnIndex"));
		int quantityColumnIndex = Integer.parseInt(props.getProperty("quantityColumnIndex"));
		int routeColumnIndex = Integer.parseInt(props.getProperty("routeColumnIndex"));


		try {
			excelFile = new File(props.getProperty("inputExcel"));
		} catch (Exception e) {
			e.printStackTrace();
		}
		XSSFWorkbook wb = new XSSFWorkbook(excelFile);
		XSSFSheet sheet = wb.getSheetAt(0);
		int ordersCounter=0;
		String uniqueOrderKey = null;
		Iterator<Row> rowIterator = sheet.rowIterator();
		Order order = null;
		OrderLines line = null;
		ArrayList<OrderLines> lines = null;
		while(rowIterator.hasNext()){
			XSSFRow row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();			
			if(row.getRowNum() != 0 && row.getCell(0) != null){
				String cellValueAsString = null;
				int cellType = 9999;
				try {
					cellType = row.getCell(0).getCellType();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}				

				switch (cellType) {
				case XSSFCell.CELL_TYPE_BOOLEAN:
					cellValueAsString = "" +row.getCell(0).getBooleanCellValue();
					break;
				case XSSFCell.CELL_TYPE_ERROR:
					cellValueAsString = row.getCell(0).getErrorCellString();
					break;
				case XSSFCell.CELL_TYPE_FORMULA:
					cellValueAsString = row.getCell(0).getCellFormula();
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					cellValueAsString = "" + row.getCell(0).getNumericCellValue();
					break;
				case XSSFCell.CELL_TYPE_STRING:
					cellValueAsString = row.getCell(0).getStringCellValue();
					break;
				default:
					cellValueAsString = "";
					break;
				}

				if(uniqueOrderKey == null){
					//This is the first order
					ordersCounter++;
					uniqueOrderKey = ordersCounter + props.getProperty("uniqueKeyJoiner") + cellValueAsString;
					order = new Order();
				}else{
					//This is not the first Order. Create a new temp key and validate
					String tempKey = ordersCounter + props.getProperty("uniqueKeyJoiner") + cellValueAsString;

					//Only create a new object if the temp key is different from the unique key
					if(!tempKey.equals(uniqueOrderKey)){
						ordersCounter++;
						uniqueOrderKey = ordersCounter + props.getProperty("uniqueKeyJoiner") + cellValueAsString;
						order.setLines(lines);
						if(orders == null) orders = new ArrayList<Order>();
						orders.add(order);
						order = new Order();
						line = null;
						lines = null;
					}
				}
			}
			while(cellIterator.hasNext()){
				XSSFCell cell = (XSSFCell) cellIterator.next();
				if(order == null) continue;
				
				String cellValueAsString = null;
				int cellType = 9999;

				try {
					cellType = cell.getCellType();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					continue;
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

				int cellNum = cell.getColumnIndex();

				if(cellNum == POColumnIndex) order.setPO(cellValueAsString);
				if(cellNum == soldToColumnIndex) order.setSoldTo(cellValueAsString);
				if(cellNum == shipToColumnIndex) order.setShipTo(cellValueAsString);
				if(cellNum == dropshipIndicatorColumnIndex) order.setDropshipIndicator(cellValueAsString);
				if(cellNum == dropshipPoColumnIndex) order.setDropshipPo(cellValueAsString);
				if(cellNum == requestedDeliveryColumnIndex) {
					if(cellType == XSSFCell.CELL_TYPE_NUMERIC){
						if(DateUtil.isCellDateFormatted(cell)){
							SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
							order.setRequestedDelivery(sdf.format(cell.getDateCellValue()));
						}
					}else{
						order.setRequestedDelivery(cellValueAsString);
					}
				}
				if(cellNum == internalNotesColumnIndex) order.setInternalNotes(cellValueAsString);
				if(cellNum == productCodeColumnIndex){
					if(line == null){
						line = new OrderLines();
					} else{
						if(lines == null){
							lines = new ArrayList<OrderLines>(); 
						}
						lines.add(line);
						line = new OrderLines();
					}
					line.setProductCode(cellValueAsString);
				}
				if(cellNum == quantityColumnIndex) line.setQuantity(cellValueAsString);
				if(cellNum == routeColumnIndex) line.setRoute(cellValueAsString);

			}
		}
		System.out.println("All Done! Closing work book");
		wb.close();
		displayExtractedOrders(orders);
	}
	
	public static void displayExtractedOrders(List<Order> orders){
		System.out.println("\n\n\n\n\n");
		System.out.println("Below is the summary of extracted orders from the excel file: \n\n");
		for(Order order:orders){
			System.out.println("Header Data: ");
			System.out.println("PO: "+order.getPO());
			System.out.println("soldTo: "+order.getSoldTo());
			System.out.println("shipTo: "+order.getShipTo());
			System.out.println("dropshipIndicator: "+order.getDropshipIndicator());
			System.out.println("requestedDelivery: "+order.getRequestedDelivery());
			System.out.println("internalNotes: "+order.getInternalNotes());
			System.out.println("\n\n");
			System.out.println("Line Data: ");
			int i=0;
			for(OrderLines line: order.getLines()){
				System.out.println("Line #"+ ++i);
				System.out.println("Product Code: "+line.getProductCode());
				System.out.println("Product Quantity: "+line.getQuantity());
				System.out.println("Route Code: "+line.getRoute());
			}
		}
	}
}
