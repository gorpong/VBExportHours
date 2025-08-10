package io.github.gorpong.hoursexport;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Properties;

/**
 * The properties file configuration controlling how the project works.
 * It used to look for the file in the class path, but once converted to a Maven
 * project, that no longer worked, so just added the ability to pass in the
 * argument on the command line and look for the default file in the current
 * working directory.  Not as useful as a full class-path search, but at least
 * this works.
 *
 * @author Gordon Galligher - gorpong@gmail.com
 */
public class ConfigProperties {
    private enum Source { NONE, FILESYSTEM, CLASSPATH }

    private Source source = Source.NONE;
    private Path loadedPath = null; // Filesystem
    private String loadedResource = null; // Classpath

    private HashMap<String, String> propList = new HashMap<String, String>();
    private String defPropFname = "VBHoursExport.properties";
    private Properties prop;

    /**
     * Create a new instance of the ConfigProperties class.
     */
    public ConfigProperties() {
        this.prop = new Properties();
    }

    /**
     * Get the property values from the specified properties file somewhere in the classpath.
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
        InputStream stream = null;

        try {
            Path path = FileSystems.getDefault().getPath(propFileName);
            if (Files.exists(path)) {
                stream = Files.newInputStream(path);
                source = Source.FILESYSTEM;
                loadedPath = path;
            } else {
                // fallback to classpath resource
                stream = getClass().getClassLoader().getResourceAsStream(propFileName);
                if (stream == null) {
                    throw new FileNotFoundException("Cannot find properties file <" + propFileName + "> in ClassPath.");
                }
                source = Source.CLASSPATH;
                loadedResource = propFileName;
            }

            prop.load(stream);
            propList.clear(); // ensure fresh
            for (Object key : prop.keySet()) {
                propList.put((String) key, (String) prop.get(key));
            }
        } catch (NoSuchFileException e) {
            throw new FileNotFoundException("Cannot find properties file <" + propFileName + "> in ClassPath.");
        } finally {
            if (stream != null) {
                stream.close();
            }
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
        if (propList.containsKey(key))
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

    /**
     * Set a configuration entry with the specified key and value.
     *
     * @param key
     * 		The key to set in the configuration database
     * @param value
     * 		The value to set for the specified key
     */
    public void setConfig(String key, String value) {
        propList.put(key, value);
        prop.setProperty(key, value);
    }

    /**
     * Save the properties to the specified properties file.
     *
     * @param propFileName Name of property file to save (defaults to defPropFname)
     * @throws IOException
     * 		Problem closing/writing file
     */
    public void saveProperties(String propFileName) throws IOException {
        Path path = FileSystems.getDefault().getPath(propFileName);
        try (FileOutputStream output = new FileOutputStream(path.toFile())) {
            prop.store(output, null);
        }
    }

    /**
     * Save the properties to the default properties file, only if it was a file.
     *
     * @throws IllegalStateException
     *      Request to save when it was found on the classpath
     * @throws IOException
     * 		Problem closing/writing file
     */
    public void saveProperties() throws IOException {
        if ( source == Source.FILESYSTEM && loadedPath != null )
            this.saveProperties(defPropFname);
        else {
            throw new IllegalStateException(
                "Properties were loaded from a classpath resource (" + loadedResource + 
                "); provide an explicit filesystem path to save to."
            );
        }
    }
}
