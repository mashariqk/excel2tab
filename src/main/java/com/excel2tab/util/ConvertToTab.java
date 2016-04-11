package com.excel2tab.util;

import java.io.File;
import java.io.InputStream;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertToTab {

	public static void main (String[] args) throws Exception {

		Properties props = new Properties();
		File excelFile =null;
		try{
			InputStream resourceStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("project.properties");
			System.out.println("resourceStream is "+resourceStream);
			props.load(resourceStream);
		}catch(Exception e){
			e.printStackTrace();
		}
		
		try {
			excelFile = new File(props.getProperty("inputExcel"));
			System.out.println("Input dir is "+props.getProperty("inputExcel"));
			System.out.println("wbStream is "+excelFile);
		} catch (Exception e) {
			e.printStackTrace();
		}
		Workbook wb = new XSSFWorkbook(excelFile);
	}
}
