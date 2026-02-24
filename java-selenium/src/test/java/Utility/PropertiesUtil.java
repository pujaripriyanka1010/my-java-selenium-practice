package Utility;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.Properties;
import java.util.Set;

public class PropertiesUtil {
	
	public static void main(String[] args) throws IOException {
		FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\Files\\config.properties");
		
		Properties prop = new Properties();
		prop.load(file);
		
		System.out.println(prop.getProperty("url"));
		System.out.println(prop.getProperty("password"));
		
		//reading all keys
		Set<String> keys = prop.stringPropertyNames();
		System.out.println(keys);
		
		Set<Object> keys2 = prop.keySet();
		System.out.println(keys2);
		
		//reading values
		Collection<Object> values = prop.values();
		System.out.println(values);
	}

}
