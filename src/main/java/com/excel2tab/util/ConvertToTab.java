package com.excel2tab.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
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

	static Properties props = null;

	public static void main (String[] args) throws Exception {

		props = new Properties();
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
		int lastRowIndex = sheet.getLastRowNum();
		int ordersCounter=0;
		String uniqueOrderKey = null;
		Iterator<Row> rowIterator = sheet.rowIterator();
		Order order = null;
		OrderLines line = null;
		ArrayList<OrderLines> lines = null;
		while(rowIterator.hasNext()){
			XSSFRow row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();	
			if(row.getRowNum() != 0 && row.getCell(0) != null && lastRowIndex != row.getRowNum()){
				String cellValueAsString = null;
				int cellType = 9999;
				try {
					cellType = row.getCell(0).getCellType();
				} catch (Exception e) {
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
						lines.add(line);
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

				if(!cellIterator.hasNext() && lastRowIndex == row.getRowNum()){
					if(lines == null) lines = new ArrayList<OrderLines>(); 
					lines.add(line);
					order.setLines(lines);
					if(orders == null) orders = new ArrayList<Order>();
					orders.add(order);
				}

			}
		}
		createTabDelimitedFile(orders);
		wb.close();
		displayExtractedOrders(orders);
	}

	public static void createTabDelimitedFile(List<Order> orders){
		StringBuffer tempBufferForFile = new StringBuffer();
		tempBufferForFile.append(instantiateFileForSAP(new StringBuffer()));
		for(Order order:orders){
			tempBufferForFile.append(headerAddendum(new StringBuffer()));
			tempBufferForFile.append(order.getPO());
			tempBufferForFile.append("\t");
			tempBufferForFile.append(order.getSoldTo());
			tempBufferForFile.append("\t");
			tempBufferForFile.append(order.getShipTo());
			tempBufferForFile.append("\t");
			tempBufferForFile.append(order.getDropshipIndicator());
			tempBufferForFile.append("\t");
			tempBufferForFile.append(order.getDropshipPo());
			tempBufferForFile.append("\t");
			tempBufferForFile.append(props.getProperty("orderReason"));
			tempBufferForFile.append("\t");
			tempBufferForFile.append(props.getProperty("POType"));
			tempBufferForFile.append("\t");
			tempBufferForFile.append(order.getRequestedDelivery());
			tempBufferForFile.append("\t");
			tempBufferForFile.append(order.getInternalNotes());
			tempBufferForFile.append("\t");
			tempBufferForFile.append(props.getProperty("fakeUserEmail"));
			tempBufferForFile.append("\n");
			for(OrderLines line: order.getLines()){
				tempBufferForFile.append(lineAddendum(new StringBuffer()));
				tempBufferForFile.append(line.getProductCode());
				tempBufferForFile.append("\t");
				tempBufferForFile.append(line.getQuantity());
				tempBufferForFile.append("\t");
				tempBufferForFile.append(props.getProperty("fakeUOM"));
				tempBufferForFile.append("\t");
				tempBufferForFile.append(line.getRoute());
				tempBufferForFile.append("\n");
			}
		}

		try (Writer writer = new BufferedWriter(new OutputStreamWriter(
				new FileOutputStream(props.getProperty("targetTabbedFile")), props.getProperty("targetTabbedCharSet")))) {
			writer.write(tempBufferForFile.toString());
		}catch(Exception e){
			e.printStackTrace();
		}

	}

	public static String instantiateFileForSAP(StringBuffer instantiationBuffer){
		instantiationBuffer.append("E");
		instantiationBuffer.append("\t");
		instantiationBuffer.append(props.getProperty("headerEmail"));
		instantiationBuffer.append("\n");
		instantiationBuffer.append("CH");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Order Type");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Salesorg");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Dis");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Div");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("PO");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Sold-to");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Ship to");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Dropship");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Drop-ship PO");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Order reason");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("PO type");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Req Deliv");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Header internal notes");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("User Email");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Product");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("Quantity");
		instantiationBuffer.append("\t");
		instantiationBuffer.append("UOM");
		instantiationBuffer.append("\t");		
		instantiationBuffer.append("Route");	
		instantiationBuffer.append("\n");	
		return instantiationBuffer.toString();
	}

	public static String headerAddendum(StringBuffer sbHeader){
		sbHeader.append("H");
		sbHeader.append("\t");
		sbHeader.append(props.getProperty("orderType"));
		sbHeader.append("\t");
		sbHeader.append(props.getProperty("salesOrg"));
		sbHeader.append("\t");
		sbHeader.append(props.getProperty("dis"));
		sbHeader.append("\t");
		sbHeader.append(props.getProperty("div"));
		sbHeader.append("\t");
		return sbHeader.toString();
	}

	public static String lineAddendum(StringBuffer sbLine){
		sbLine.append("I");
		for(int i=0; i < Integer.parseInt(props.getProperty("tabsAtLineStart")); i++) sbLine.append("\t");
		return sbLine.toString();
	}

	public static void displayExtractedOrders(List<Order> orders){
		System.out.println("\n\n\n\n\n");
		System.out.println("Below is the summary of extracted orders from the excel file: \n\n");
		int i=0;
		for(Order order:orders){
			System.out.println("\n\n");
			System.out.println("Order #"+ ++i);
			System.out.println("Header Data: ");
			System.out.println("PO: "+order.getPO());
			System.out.println("soldTo: "+order.getSoldTo());
			System.out.println("shipTo: "+order.getShipTo());
			System.out.println("dropshipIndicator: "+order.getDropshipIndicator());
			System.out.println("requestedDelivery: "+order.getRequestedDelivery());
			System.out.println("internalNotes: "+order.getInternalNotes());
			System.out.println("\n\n");
			System.out.println("Line Data: ");
			int j=0;
			for(OrderLines line: order.getLines()){
				System.out.println("Line #"+ ++j);
				System.out.println("Product Code: "+line.getProductCode());
				System.out.println("Product Quantity: "+line.getQuantity());
				System.out.println("Route Code: "+line.getRoute());
			}
		}
	}
}
