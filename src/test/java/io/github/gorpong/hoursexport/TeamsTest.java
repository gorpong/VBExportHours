package io.github.gorpong.hoursexport;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import static org.junit.jupiter.api.Assertions.assertEquals;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.ArgumentMatchers.eq;
import org.mockito.Mock;
import org.mockito.MockedStatic;
import static org.mockito.Mockito.mockStatic;
import static org.mockito.Mockito.when;
import org.mockito.junit.jupiter.MockitoExtension;

@ExtendWith(MockitoExtension.class)
class TeamsTest {

    /*
     * Create a logger to log each test as it runs.
    */
    private static final Logger logger = Logger.getLogger(TeamsTest.class.getName());
    private static boolean setupPrinted = false;

    static {
            ConsoleHandler handler = new ConsoleHandler();
            handler.setFormatter(new SimpleFormatterWithoutDate());
            logger.addHandler(handler);
            logger.setUseParentHandlers(false);
    }

    @Mock
    private ConfigProperties config;

    private Teams teams;
    private Workbook workbook;
    private Sheet sheet;
    private MockedStatic<WorkbookFactory> workbookFactoryMock;

    @BeforeEach
    void setUp() {
	    if (!setupPrinted) {
    		logger.log(Level.INFO, String.format("%s: Setting up test environment", this.getClass().getName()));
		setupPrinted = true;
	    }

        // Clear out the Student's list each time
        Student.clearStudents();
        
        // Create a mock workbook and sheet
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("TestSheet");

        // Mock ConfigProperties
        when(config.getConfig(eq("inputColumnHours"), anyString())).thenReturn("workday_w");
        when(config.getConfig(eq("inputColumnName"), anyString())).thenReturn("Name");
        when(config.getConfig(eq("inputColumnID"), anyString())).thenReturn("empno");
        when(config.getConfig(eq("inputColumnTeam"), anyString())).thenReturn("Department");

        // Initialize Teams with a mock file name and config
        teams = new Teams("mockFileName", config);

        // Mock WorkbookFactory.create
        workbookFactoryMock = mockStatic(WorkbookFactory.class);
        workbookFactoryMock.when(() -> WorkbookFactory.create(any(File.class))).thenReturn(workbook);
    }

    // Refactor the creation of the mock spreadsheet for each test
    void createMockSheetData(String team, String name, String id, Integer hours) {
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Department");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("empno");
        headerRow.createCell(3).setCellValue("workday_w");

        Row dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
        dataRow.createCell(0).setCellValue(team);
        dataRow.createCell(1).setCellValue(name);
        dataRow.createCell(2).setCellValue(id);
        dataRow.createCell(3).setCellValue(hours);
    }

    @Test
    void testParseExcel() throws IOException, InvalidFormatException {
        logger.log(Level.INFO, String.format("Running testParseExcel, Size of Teams:  %d", teams.getTeams().size()));

        // Create mock data in the sheet
        createMockSheetData("1", "John Doe", "123", 40);
        createMockSheetData("1", "Jane Doe", "456", 35);

        // Parse Excel
        teams.parseExcel();

        // Verify the team data
        Map<Integer, List<Student>> teamData = teams.getTeamsData();
        assertEquals(1, teamData.size());
        List<Student> students = teamData.get(1);
        assertEquals(2, students.size());
        assertEquals("Doe, John", students.get(0).getName());
        assertEquals("Doe, Jane", students.get(1).getName());
    }

    @Test
    void testGetTeams() throws IOException, EncryptedDocumentException, InvalidFormatException {
        logger.log(Level.INFO, String.format("Running testGetTeams, Size of Teams:  %d", teams.getTeams().size()));

        // Create mock data in the sheet
        createMockSheetData("1", "John Doe", "123", 40);
        createMockSheetData("2", "Jane Doe", "456", 35);

        // Parse Excel
        teams.parseExcel();

        // Get teams
        List<Integer> teamsList = teams.getTeams();
        assertEquals(2, teamsList.size());
        assertEquals(1, teamsList.get(0).intValue());
        assertEquals(2, teamsList.get(1).intValue());
    }

    @Test
    void testGetHoursByTeam() throws IOException, EncryptedDocumentException, InvalidFormatException {
        logger.log(Level.INFO, String.format("Running testGetHoursByTeam, Size of Teams:  %d", teams.getTeams().size()));

        // Create mock data in the sheet
        createMockSheetData("1", "John Doe", "123", 40);
        createMockSheetData("1", "Jane Doe", "456", 35);

        // Parse Excel
        teams.parseExcel();

        // Get hours by team
        List<Student> students = teams.getHoursByTeam(1);
        assertEquals(2, students.size());
        assertEquals("Doe, John", students.get(0).getName());
        assertEquals(40, students.get(0).getHours(), 0.001);
        assertEquals("Doe, Jane", students.get(1).getName());
        assertEquals(35, students.get(1).getHours(), 0.001);
    }

    @Test
    void testFindColumn() {
        logger.log(Level.INFO, String.format("Running testFindColumn, Size of Teams:  %d", teams.getTeams().size()));

        // Create mock data in the sheet
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Department");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("empno");
        headerRow.createCell(3).setCellValue("workday_w");

        // Find column
        int colTeam = teams.findColumn(sheet, "Department");
        assertEquals(0, colTeam);

        int colName = teams.findColumn(sheet, "Name");
        assertEquals(1, colName);

        int colID = teams.findColumn(sheet, "empno");
        assertEquals(2, colID);

        int colHours = teams.findColumn(sheet, "workday_w");
        assertEquals(3, colHours);
    }

    @AfterEach
    void cleanUp() {
        if (workbookFactoryMock != null) {
            workbookFactoryMock.close();
        }
    }
}
