package com.excel2tab.util;

/**
 * 
 * @author Mashariq Khan
 * POJO file to hold Order line data parsed from the excel file
 *
 */
public class OrderLines {
	
	private String productCode;
	private String quantity;
	private String route;

	public String getProductCode() {
		return productCode;
	}
	public void setProductCode(String productCode) {
		this.productCode = productCode;
	}
	public String getQuantity() {
		return quantity;
	}
	public void setQuantity(String quantity) {
		this.quantity = quantity;
	}
	public String getRoute() {
		return route;
	}
	public void setRoute(String route) {
		this.route = route;
	}
	
}
