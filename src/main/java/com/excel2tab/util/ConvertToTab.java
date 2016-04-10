package com.excel2tab.util;

import java.io.InputStream;
import java.util.Properties;

public class ConvertToTab {

	public static void main(String[] args){

		Properties props = new Properties();
		try{
			InputStream resourceStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("project.properties");
			props.load(resourceStream);
			System.out.println("booyah bru");
			System.out.println(props.getProperty("hola"));
		}catch(Exception e){
			e.printStackTrace();
		}
	}
}
