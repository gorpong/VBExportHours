package io.github.gorpong.hoursexport;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

@DisplayName("Properties File Tests")
public class ConfigPropertiesTest {
    private ConfigProperties config;
    private final String TEST_PROP_FILE = "test.properties";

    @BeforeEach
    public void setUp() {
        config = new ConfigProperties();
        // Create a sample properties file for testing
        try {
            java.nio.file.Files.write(new File(TEST_PROP_FILE).toPath(),
                    ("key1=value1\nkey2=value2\nkey3=value3").getBytes());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @AfterEach
    public void tearDown() {
        // Clean up the test properties file
        new File(TEST_PROP_FILE).delete();
    }

    @Nested
    @DisplayName("Specific File Tests")
    class SpecificConfigFile {

        @Test
        public void testLoadSpecificFile() {
            try {
                config.getPropValues(TEST_PROP_FILE);
                assertEquals("value1", config.getConfig("key1"));
                assertEquals("value2", config.getConfig("key2"));
                assertEquals("value3", config.getConfig("key3"));
            } catch (IOException e) {
                fail("Unexpected exception: " + e.getMessage());
            }
        }

        @Test
        public void testExistingKey() {
            try {
                config.getPropValues(TEST_PROP_FILE);
                assertEquals("value1", config.getConfig("key1"));
            } catch (IOException e) {
                fail("Unexpected exception: " + e.getMessage());
            }
        }

        @Test
        public void testSetConfig() throws IOException {
            config.getPropValues(TEST_PROP_FILE);
            config.setConfig("key4", "value4");
            config.setConfig("key1", "newValue1");

            assertEquals("value4", config.getConfig("key4"));
            assertEquals("newValue1", config.getConfig("key1"));
        }

        @Test
        public void testSaveProperties() throws IOException {
            config.getPropValues(TEST_PROP_FILE);
            config.setConfig("key4", "value4");
            config.setConfig("key1", "newValue1");

            config.saveProperties(TEST_PROP_FILE);

            ConfigProperties newConfig = new ConfigProperties();
            newConfig.getPropValues(TEST_PROP_FILE);

            assertEquals("value4", newConfig.getConfig("key4"));
            assertEquals("newValue1", newConfig.getConfig("key1"));
        }
    }

    @Nested
    @DisplayName("Load the Default Classpath-search Properties file")
    class ClassPathSearch {

        @Test
        public void testLoadDefaultFile() {
            try {
                config.getPropValues();
                assertEquals("value1", config.getConfig("key1"));
                assertEquals("value2", config.getConfig("key2"));
                assertEquals("value3", config.getConfig("key3"));
            } catch (IOException e) {
                fail("Unexpected exception: " + e.getMessage());
            }
        }

        @Test
        public void testSaveDefaultFile() {
            try {
                config.getPropValues();
                config.setConfig("key4", "value4");
                assertThrows(
                    IllegalStateException.class, 
                    () -> config.saveProperties()
                );
            } catch (Exception e) {
                fail("Unexpected exception: " + e.getMessage());
            }
        }
    }

    @Nested
    @DisplayName("Edge Cases for Properties File")
    class PropertiesEdgeCases {
        @Test
        public void testFileNotFound() {
            ConfigProperties config = new ConfigProperties();
            Exception exception = assertThrows(FileNotFoundException.class, () -> {
                config.getPropValues("nonexistent.properties");
            });
            String expectedMessage = "Cannot find properties file <nonexistent.properties> in ClassPath.";
            String actualMessage = exception.getMessage();
            assertTrue(actualMessage.contains(expectedMessage));
        }

        @Test
        public void testNonExistingKey() {
            try {
                config.getPropValues(TEST_PROP_FILE);
                assertEquals("default", config.getConfig("nonexistentKey", "default"));
            } catch (IOException e) {
                fail("Unexpected exception: " + e.getMessage());
            }
        }
    }
}
