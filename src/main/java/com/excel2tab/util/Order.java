package com.excel2tab.util;

import java.util.List;

/**
 * 
 * @author Mashariq Khan
 * POJO file to hold Order header data parsed from the excel file
 *
 */
public class Order {

	private String PO;
	private String soldTo;
	private String shipTo;
	private String dropshipIndicator;
	private String dropshipPo;	
	private String requestedDelivery;
	private String internalNotes;
	private List<OrderLines> lines;
	
	public String getPO() {
		return PO;
	}
	public void setPO(String pO) {
		PO = pO;
	}
	public String getSoldTo() {
		return soldTo;
	}
	public void setSoldTo(String soldTo) {
		this.soldTo = soldTo;
	}
	public String getShipTo() {
		return shipTo;
	}
	public void setShipTo(String shipTo) {
		this.shipTo = shipTo;
	}
	public String getDropshipIndicator() {
		return dropshipIndicator;
	}
	public void setDropshipIndicator(String dropshipIndicator) {
		this.dropshipIndicator = dropshipIndicator;
	}
	public String getDropshipPo() {
		return dropshipPo;
	}
	public void setDropshipPo(String dropshipPo) {
		this.dropshipPo = dropshipPo;
	}
	public String getRequestedDelivery() {
		return requestedDelivery;
	}
	public void setRequestedDelivery(String requestedDelivery) {
		this.requestedDelivery = requestedDelivery;
	}
	public String getInternalNotes() {
		return internalNotes;
	}
	public void setInternalNotes(String internalNotes) {
		this.internalNotes = internalNotes;
	}
	public List<OrderLines> getLines() {
		return lines;
	}
	public void setLines(List<OrderLines> lines) {
		this.lines = lines;
	}
}
