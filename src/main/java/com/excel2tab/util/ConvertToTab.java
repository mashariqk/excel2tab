package com.excel2tab.util;

import java.io.File;
import java.io.InputStream;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertToTab {

	public static void main (String[] args) throws Exception {

		Properties props = new Properties();
		File excelFile =null;
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
	}
}
