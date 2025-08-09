package io.github.gorpong.hoursexport;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.util.ArrayList;
import java.util.Collections;
import java.io.File;
import java.io.IOException;

/**
 * Scan the Excel hours class export file and build structure of People to Teams.
 * 
 * This uses the Apache POI (v3.17) for all Excel related processing.
 * TODO Should really move all the Excel-related stuff to its own class
 * 
 * @author Gordon Galligher - gorpong@gmail.com
 */
public class Teams {
	private HashMap<Integer, ArrayList<Student>> _teams = new HashMap<>();
	private String fileName;
	private ConfigProperties config;
	
	private String columnHours = null; private final String defColHrs  = "workday_w";
	private String columnName  = null; private final String defColName = "Name";
	private String columnID    = null; private final String defColID   = "empno";
	private String columnTeam  = null; private final String defColTeam = "Department";
	
	/**
	 * Construct the Teams object for parsing the {@code file} to get hours.
	 * 
	 * @param file	
	 * 		The file to parse (in Excel format)
	 * @param cfg	
	 * 		The configuration object for grabbing things like column name
	 */
	public Teams(String file, ConfigProperties cfg) {
		fileName = file;
		config   = cfg;
		columnHours = config.getConfig("inputColumnHours", defColHrs);
		columnName  = config.getConfig("inputColumnName", defColName);
		columnID    = config.getConfig("inputColumnID", defColID);
		columnTeam  = config.getConfig("inputColumnTeam", defColTeam);
	}
	
	/**
	 * Parse the Excel file for this instance and create the data structure holding the information.
	 * 
	 * @throws IOException
	 * 			Error when closing workbook 
	 * @throws InvalidFormatException
	 *  		Error when creating workbook based on file
	 * @throws EncryptedDocumentException
	 * 		 	Workbook in file is encrypted
	 * @throws IllegalStateException
	 * 			Workbook doesn't have columns we're looking to find
	 * 
	 */
	public void parseExcel() throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook workbook = WorkbookFactory.create(new File(fileName));
		Sheet sheet = workbook.getSheetAt(0);

		int colName  = findColumn(sheet, columnName);
		int colHours = findColumn(sheet, columnHours);
		int colID    = findColumn(sheet, columnID);
		int colTeam  = findColumn(sheet, columnTeam);
		if ( colName < 0 || colHours < 0 || colID < 0 || colTeam < 0 ) {
			String msg = "Can't find appropriate columns, missing:  ";
			if ( colName < 0 ) msg += columnName + " ";
			if ( colHours <0 ) msg += columnHours + " ";
			if ( colID < 0 )   msg += columnID + " ";
			if ( colTeam < 0 ) msg += columnTeam;
			throw new IllegalStateException(msg);
		}
		Integer team = 0;
		Student student;
		String tm = null;
		for ( Row row : sheet ) {
			if ( row.getRowNum() == 0 )
				continue;
			try {
				tm = row.getCell(colTeam).getStringCellValue();
				team = Integer.parseInt(tm);
			} catch (Exception e) {
				System.err.println("Parse Error:  row "+row.getRowNum()+", invalid Team number:  " + tm);
				continue;
			}
			try {
				student = Student.getStudent(row.getCell(colName).getStringCellValue(),
						row.getCell(colID).getStringCellValue(),
						row.getCell(colHours).getNumericCellValue());
				if ( _teams.containsKey(team) ) {
					if ( ! _teams.get(team).contains(student) )
						_teams.get(team).add(student);
				} else {
					ArrayList<Student> ppl = new ArrayList<Student>();
					ppl.add(student);
					_teams.put(team, ppl);
				}
			} catch (IllegalStateException e) {
				System.err.println("Cannot parse row " + row.getRowNum() + " to get appropriate data");
			}
		}
		workbook.close();
	}

	/**
	 * Get the list of teams that we've parsed as an {@code ArrayList<Integer>}.
	 * 
	 * @return
	 * 		The list of sorted teams
	 */
	public ArrayList<Integer> getTeams() {
		ArrayList<Integer> teams = new ArrayList<Integer>();
		teams.addAll(_teams.keySet());
		Collections.sort(teams);
		return teams;
	}
	
	/**
	 * Get the list of {@code Students} for a specific team, sorted by the number
	 * of hours worked in descending order (e.g., most hours to least).
	 * 
	 * @param team
	 * 		The team from which to get the sorted list of {@code Students}
	 * @return
	 * 		The sorted list of {@code Students} (or null if error)
	 */
	public ArrayList<Student> getHoursByTeam(int team) {
		try {
			ArrayList<Student> sortedList = _teams.get(team);
			sortedList.sort((a, b) -> a.compareTo(b));
			return sortedList;
		} catch (NullPointerException e) {
			return null;
		}
	}

	/**
	 * Print the team information, a debug method, not for production use.
	 */
	protected void printTeamInfo() {
		ArrayList<Integer> teams = getTeams();
		for (Integer team : teams) {
			System.out.println("" + team + ":  " + getHoursByTeam(team));
		}
	}

	public Map<Integer, List<Student>> getTeamsData() {
		return new HashMap<>(_teams); 
	}

	/**
	 * Find the specific position of the column containing the header label we want.
	 * 
	 * @param sheet
	 * 		The sheet to look in
	 * @param colSearch
	 * 		The string of the column we're looking for (-1 if not found)
	 * @return
	 * 		The column number where the column is found
	 */
	public int findColumn(Sheet sheet, String colSearch) {
       		 return findColumnInternal(sheet, colSearch);
	}
    
	private int findColumnInternal(Sheet sheet, String colSearch) {
		int column = -1;	
		
		Row row0 = sheet.getRow(0);
		for ( Cell cell : row0 ) {
			if ( cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().equals(colSearch) ) {
				column = cell.getColumnIndex();
				break;
			}
		}
		return column;
	}

}
