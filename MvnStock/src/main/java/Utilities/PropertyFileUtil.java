package Utilities;

import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.util.Properties;

public class PropertyFileUtil 
{
	public static String getValueForKey(String key) throws Throwable 
	{
		Properties configproperties=new Properties();
		FileInputStream fis=new FileInputStream
	("D:\\ESTHAN\\AUTOMATION\\Automation Testing\\MvnStock\\PropertyFiles\\Environment.properties");
		configproperties.load(fis);
		
		return configproperties.getProperty(key);
		
	}
	
}

