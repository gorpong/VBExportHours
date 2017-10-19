package net.gginbox.hoursExport;

/**
 * Read our configuration file that has the information for how to scan
 * and parse the Excel file as well as how to build the new one.
 */

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Properties;

/**
 * The properties file configuration controlling how the project works.
 * It used to look for the file in the class path, but once converted to a Maven
 * project, that no longer worked, so just added the ability to pass in the
 * argument on the command line and look for the default file in the current
 * working directory.  Not as useful as a full class-path search, but at lesat
 * this works.
 * 
 * @author Gordon Galligher - gorpong@gginbox.net
 *
 */
public class ConfigProperties {
	private HashMap<String, String> propList = new HashMap<String,String>();
	private String defPropFname = "VBHoursExport.properties";

	/**
	 * Get the property values from the VBHoursExport.properties file somewhere in the classpath.
	 * Read through all the entries and build a private HashMap of key/value pairs.  This
	 * can then be used in the {@code getConfig} method to get at the actual value associated
	 * with that information.
	 * 
	 * @param propFileName Name of property file to parse (defaults to defPropFname)
	 *  
	 * @throws IOException
	 * 		Problem closing/writing file
	 * @throws FileNotFoundException
	 * 		Can't find properties file
	 */
	public void getPropValues(String propFileName) throws IOException, FileNotFoundException {
		Properties prop = new Properties();
		Path path = FileSystems.getDefault().getPath(propFileName);
		InputStream stream = null;
		
		try {
//			stream = ClassLoader.getSystemClassLoader().getResourceAsStream(propFileName);
			stream = Files.newInputStream(path);
			if ( stream != null ) {
				prop.load(stream);
			} else {
				throw new FileNotFoundException("Cannot find properties file <"+propFileName+"> in ClassPath.");
			}
			for (Object key : prop.keySet()) {
				propList.put((String) key, (String) prop.get(key));
			}
		} catch (Exception e) {
			System.out.println("Exception:  " + e);
		} finally {
			stream.close();
		}
	}
	
	/**
	 * Get the property values from the properties file that is the default
	 * for the program.  Really, just call the other constructor with the
	 * default file name attribute.
	 * 
	 * @throws IOException
	 * 		Problem closing/writing file
	 * @throws FileNotFoundException
	 * 	 	Can't find properties file
	 */
	public void getPropValues() throws IOException, FileNotFoundException {
		this.getPropValues(defPropFname);
	}

	/**
	 * Get the configuration entry pointed at by {@code key}.
	 * 
	 * @param key
	 * 		The key to lookup in the configuration database
	 * @return
	 * 		The value associated with that key (or null if not found)
	 */
	public String getConfig(String key) {
		if ( propList.containsKey(key) )
			return propList.get(key);
		else return null;
	}
	
	/**
	 * Get the configuration entry pointed at by {@code key}, or {@code default} if not present.
	 * 
	 * @param key
	 * 		The key to lookup in the configuration database
	 * @param defVal
	 * 		The default value to return, if {@code key} isn't found.  
	 * @return
	 * 		The value associated with that key (or default)
	 */
	public String getConfig(String key, String defVal) {
		String value = getConfig(key);
		return value != null ? value : defVal;
	}
	
}
