package net.gginbox.hoursExport;

import java.util.logging.ConsoleHandler;
import java.util.logging.LogRecord;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class SimpleFormatterWithoutDate extends SimpleFormatter {
    @Override
    public String format(LogRecord record) {
    return String.format("%s: %s\n", record.getLevel(), record.getMessage());
    }
}

class StudentTest {

    private static final Logger logger = Logger.getLogger(StudentTest.class.getName());
    private static boolean setupPrinted = false;

    static {
	    ConsoleHandler handler = new ConsoleHandler();
	    handler.setFormatter(new SimpleFormatterWithoutDate());
	    logger.addHandler(handler);
	    logger.setUseParentHandlers(false);
    }

    @BeforeEach
    void setUp() {
        if (!setupPrinted) {
            logger.info("Setting up test environment");
            setupPrinted = true;
        }
        Student.clearStudents();
    }

    @Test
    void testGetStudent_NewStudent() {
        logger.info("Running testGetStudent_NewStudent");
        Student person = Student.getStudent("Joe Bob", "xyz123", 5.5);
        assertNotNull(person);
        assertEquals("Bob", person.getName().split(",")[0]);
        assertEquals("xyz123", person.getId());
        assertEquals(5.5, person.getHours(), 0.01);
    }

    @Test
    void testGetStudent_ExistingStudent() {
        logger.info("Running testGetStudent_ExistingStudent");
        Student person1 = Student.getStudent("Joe Bob", "xyz123", 5.5);
        Student person2 = Student.getStudent("Joe Bob", "xyz123", 2.5);
        assertEquals(person1, person2);
        assertEquals(8.0, person2.getHours(), 0.01);
    }

    @Test
    void testCompareTo() {
        logger.info("Running testCompareTo");
        Student person1 = Student.getStudent("Joe Bob", "xyz123", 10.0);
        Student person2 = Student.getStudent("Jane Doe", "abc456", 5.0);
        Student person3 = Student.getStudent("John Smith", "def789", 10.0);

        assertEquals(-1, person1.compareTo(person2));
        assertEquals(0, person1.compareTo(person3));
        assertEquals(1, person2.compareTo(person1));
    }

    @Test
    void testToString() {
        logger.info("Running testToString");
        Student person = Student.getStudent("Bob, Joe", "xyz123", 5.5);
        assertEquals("[Joe Bob:xyz123:5.5]", person.toString());
    }
}
