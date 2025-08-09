package io.github.gorpong.hoursexport;

import java.lang.String;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Output class for creating the output file in Excel format. Uses the Team
 * class/object for creating the output file.
 * 
 * Currently uses the Apache POI (v3.17) for all of the Excel-related stuff.
 * 
 * @author Gordon Galligher - gorpong@gmail.com
 *
 */
public class Output {

	public enum SheetType {
		COACHES, PARENTS,
	}

	/*
	 * Class variables
	 */
	private static Map<String, CellStyle> styles = null;		// Set the first time in initialize() only
	
	/*
	 * Instance variables
	 */
	private String fileName;
	private Workbook workbook;
	private Calendar date;
	private String datestr;
	private ConfigProperties config;
	private double lowHours = 0.0;
	private double highHours = 0.0;
	private int countLow = 0;		// Counter for providing stats on # too low hours
	private int countHigh = 0;		// Ditto for # of too high hours
	private List<String> highLowList; 

	/**
	 * Private constructor, use the initialize() factory method to create and start
	 * generating an output Excel file.
	 * 
	 * @param file
	 *            The output file to create
	 * @param cfg
	 *            The configuration object to grab values later
	 */
	private Output(String file, ConfigProperties cfg) {
		fileName = file;
		config = cfg;
	}

	/**
	 * Initialize the workbook {@code fname} and create it and supplemental bits
	 * we need later.  This also sets up the date values for inclusion in the Excel
	 * file and finalizes the configuration of the fonts and formats for the various
	 * cell components.
	 * 
	 * @param fname
	 * 		The output file name (the new Excel file to create)
	 * @param cfg
	 * 		The configuration object (in case we need configuration bits)
	 * @param low
	 * 		The low water mark for hours that are too low
	 * @param hi
	 * 		The high water mark for hours that are too high
	 * @return Output
	 * 		Factory method, creates new {@code Output} object and returns it
	 */
	public static Output initialize(String fname, ConfigProperties cfg, double low, double hi) {
		Output out = new Output(fname, cfg);
		out.lowHours = low;
		out.highHours = hi;
		out.highLowList = new ArrayList<String>();
		Pattern regexXLS = Pattern.compile("^.*.xls$");

		if ( regexXLS.matcher(fname).matches() ) out.workbook = new HSSFWorkbook();
		else out.workbook = new XSSFWorkbook();
		if ( styles == null ) styles = createStyles(out.workbook);
		
		out.date = Calendar.getInstance();
		out.datestr = String.format("%02d/%02d/%04d %02d:%02d %s",  out.date.get(Calendar.MONTH)+1, 
				out.date.get(Calendar.DATE), out.date.get(Calendar.YEAR), out.date.get(Calendar.HOUR), 
				out.date.get(Calendar.MINUTE), out.date.get(Calendar.AM_PM) == 1 ? "PM" : "AM");
		return out;
	}

	/**
	 * Write out and close the Excel file. Throws exceptions if there is an error
	 * writing or formatting the sheet.
	 * 
	 * @throws IOException
	 * 		Error writing/closing file
	 * 
	 */
	public void close() throws IOException {
		FileOutputStream out = new FileOutputStream(this.fileName);
		workbook.write(out);
		out.close();
		workbook.close();
	}
	
	public int getLowCount() {
		return this.countLow;
	}
	public int getHighCount() {
		return this.countHigh;
	}
	
	/**
	 * Fill the sheet with appropriate cells based on the team information and
	 * return the last row that was written into.
	 * 
	 * @param sheet
	 *            The existing Sheet object
	 * @param students
	 *            The list of people for that team (ignore sheet if null)
	 * @param team
	 *            The number of the team we are adding to the sheet
	 * @param type
	 *            The type of sheet this is (COACHES/PARENTS)
	 * @param rowStart
	 *            The starting row for the hours
	 * @param colStart
	 *            The starting column for the hours
	 * @return The ending row we've added into the sheet
	 */
	public int fillSheet(Sheet sheet, ArrayList<Student> students, Integer team, SheetType type, int rowStart, int colStart)  {
		int row = rowStart;
		int col = colStart;
		double hrs;

		if ( students == null )		// Team doesn't exist, so don't do anything 
			return rowStart;
		
		for (Student p : students) {
			col = colStart;
			String nameOrID = type == SheetType.COACHES ? p.getName() : p.getId();
			Row sheetRow = sheet.getRow(row);
			if ( sheetRow == null ) sheetRow = sheet.createRow(row);
			Cell cell = sheetRow.createCell(col++);
			cell.setCellStyle(type == SheetType.PARENTS 
					? styles.get("cell_normal_centered") 
					: styles.get("cell_normal"));
			cell.setCellValue(nameOrID);
			
			if ( type == SheetType.PARENTS ) {
				cell = sheetRow.createCell(col++);
				cell.setCellStyle(styles.get("cell_normal_centered"));
				cell.setCellValue(team);
			}
			cell = sheetRow.createCell(col++);
			hrs = p.getHours();
			if (hrs > this.highHours) {
				if ( ! this.highLowList.contains(p.getId()) )
				{
					this.countHigh++;
					highLowList.add(p.getId());
				}
				cell.setCellStyle(styles.get("cell_highlight_right"));
			} else if (hrs < this.lowHours) {
				if ( ! this.highLowList.contains(p.getId()) ) {
					this.countLow++;
					highLowList.add(p.getId());
				}
				cell.setCellStyle(styles.get("cell_bold_red_right"));
			} else {
				cell.setCellStyle(styles.get("cell_normal_right"));
			}
			cell.setCellValue(hrs);
			row++;
		}
		return row;
	}

	/**
	 * Create the worksheet passed as type and put the data into it.
	 * 
	 * @param teams
	 * 		The list of teams
	 * @param type
	 * 		The sheet type we are to create based on enum
	 */
	public void createSheet(Teams teams, SheetType type) {
		Sheet sheet;
		int maxRow = 0;
		int row;
		ArrayList<Integer> columns = new ArrayList<Integer>();

		if (type == SheetType.COACHES) {
			sheet = workbook.createSheet("Coaches");
			sheet.setDisplayGridlines(false);
			sheet.setPrintGridlines(false);
			sheet.setFitToPage(true);
			sheet.setHorizontallyCenter(true);
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));	// Title for sheet
			sheet.getPrintSetup().setLandscape(false);
			Row headerRow = sheet.createRow(0);
			headerRow.setHeightInPoints(30.60f);
			Cell cell = headerRow.createCell(0);
			cell.setCellValue("COACHES Hours Report " + datestr);
			cell.setCellStyle(styles.get("header"));

			String val = config.getConfig("coachesStartRow");
			maxRow = (val == null) ? maxRow - 1 : Integer.parseInt(val);
			for (String section : "TopRow,MidRow,BotRow".split(",")) {
				String value = config.getConfig("coaches" + section);
				if (value != null) {
					int startRow = maxRow + 1;
					int startCol;
					Row sheetRow = sheet.createRow(startRow);
					for (String team : value.split(",")) {
						try {
							startCol = Integer.parseInt(config.getConfig("coachesColumn-" + team));
						} catch (NumberFormatException e) {
							System.err.println("Config Error:  No coachesColumn-" + team + " line found");
							return;
						}
						if ( ! columns.contains(startCol) ) {
							columns.add(startCol);
							columns.add(startCol + 1);
						}
						row = fillSheet(sheet, teams.getHoursByTeam(Integer.parseInt(team)), Integer.parseInt(team),
								type, startRow + 1, startCol);
						if ( row > startRow + 1 ) {		// Put headers in, if necessary
							cell = sheetRow.createCell(startCol);
							cell.setCellStyle(styles.get("cell_normal_title_grey40"));
							cell.setCellValue("Team " + team);
							cell = sheetRow.createCell(startCol + 1);
							cell.setCellStyle(styles.get("cell_normal_title_grey40"));;
							cell.setCellValue("Hours");
							if (row > maxRow)
								maxRow = row;
						}
					}
				}
			}
			for (Integer i : columns) {
				sheet.autoSizeColumn(i);
			}
		} else {
			sheet = workbook.createSheet("Parents");
			sheet.setDisplayGridlines(false);
			sheet.setPrintGridlines(false);
			sheet.setFitToPage(true);
			sheet.setHorizontallyCenter(true);
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 12));
			sheet.getPrintSetup().setLandscape(false);
			Row headerRow = sheet.createRow(0);
			headerRow.setHeightInPoints(30.60f);
			Cell cell = headerRow.createCell(0);
			cell.setCellValue("PARENTS Hours Report " + datestr);
			cell.setCellStyle(styles.get("header"));

			String val = config.getConfig("parentsStartRow");
			maxRow = (val == null) ? maxRow - 1 : Integer.parseInt(val);
			for (String section : "TopRow,MidRow,BotRow".split(",")) {
				String value = config.getConfig("parents" + section);
				if (value != null) {
					int startRow = maxRow + 1;
					Row sheetRow = sheet.createRow(startRow);
					for (String team : value.split(",")) {
						int startCol;
						try {
							startCol = Integer.parseInt(config.getConfig("parentsColumn-" + team));
						} catch (NumberFormatException e) {
							System.err.println("Config Error:  No parentsColumn-" + team + " line found.");
							return;
						}
						if ( ! columns.contains(startCol) ) {
							columns.add(startCol);
							columns.add(startCol + 1);
						}
						row = fillSheet(sheet, teams.getHoursByTeam(Integer.parseInt(team)), Integer.parseInt(team),
								type, startRow + 1, startCol);
						if ( row > startRow + 1 ) {			// Put headers if there was data
							cell = sheetRow.createCell(startCol);
							cell.setCellStyle(styles.get("cell_normal_title_grey40"));
							cell.setCellValue("ID");
							cell = sheetRow.createCell(startCol + 1);
							cell.setCellStyle(styles.get("cell_normal_title_grey40"));
							cell.setCellValue("Team");
							cell = sheetRow.createCell(startCol + 2);
							cell.setCellStyle(styles.get("cell_normal_title_grey40"));
							cell.setCellValue("Hours");
							if (row > maxRow)
								maxRow = row;
						}
					}
				}
			}
			for (Integer i : columns) {
				sheet.autoSizeColumn(i);;
			}
		}
	}
	
	/**
	 * Create a mapping of certain styles to make it easier to use those styles.
	 * This was taken from the example code for the Apache POI project and then 
	 * modified as needed for this project.
	 * 
	 * @param wb
	 * 		The workbook object
	 * @return
	 * 		The mapping of styles
	 */
	private static Map<String, CellStyle> createStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<>();
		DataFormat df = wb.createDataFormat();
		CellStyle style;
		
		/*
		 * First the fonts, then the styles that use those fonts
		 */
		Font headerFont = wb.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 24);
		
		Font boldFont = wb.createFont();
		boldFont.setBold(true);
		
		Font boldBlueFont = wb.createFont();
		boldBlueFont.setColor(IndexedColors.BLUE.getIndex());
		boldBlueFont.setBold(true);
		
		Font boldBlue14ptFont = wb.createFont();
		boldBlue14ptFont.setFontHeightInPoints((short)14);
		boldBlue14ptFont.setColor(IndexedColors.DARK_BLUE.getIndex());
		boldBlue14ptFont.setBold(true);
		
		Font boldRedFont = wb.createFont();
		boldRedFont.setColor(IndexedColors.RED.getIndex());
		boldRedFont.setBold(true);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setIndention((short) 3);
		style.setFont(headerFont);
		styles.put("header", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(headerFont);
		style.setDataFormat(df.getFormat("d-mmm"));
		styles.put("header_date", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(boldFont);
		styles.put("cell_b", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(boldFont);
		styles.put("cell_b_centered", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(boldFont);
		style.setDataFormat(df.getFormat("d-mmm"));
		styles.put("cell_b_date", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(boldFont);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setDataFormat(df.getFormat("d-mmm"));
		styles.put("cell_g", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(boldBlueFont);
		styles.put("cell_bb", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(boldFont);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setDataFormat(df.getFormat("d-mmm"));
		styles.put("cell_bg", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(boldBlue14ptFont);
		style.setWrapText(true);
		styles.put("cell_h", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setWrapText(false);
		style.setShrinkToFit(false);
		styles.put("cell_normal", style);
		
		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(boldFont);
		style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styles.put("cell_normal_title_grey40", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setWrapText(false);
		style.setShrinkToFit(false);
		styles.put("cell_normal_centered", style);
		
		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setWrapText(false);
		style.setShrinkToFit(false);
		styles.put("cell_normal_right", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setWrapText(true);
		style.setDataFormat(df.getFormat("d-mmm"));
		styles.put("cell_normal_date", style);

		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setIndention((short)1);
		style.setWrapText(true);
		styles.put("cell_indented", style);

		style = createBorderedStyle(wb);
		style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styles.put("cell_blue", style);
		
		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(boldRedFont);
		styles.put("cell_bold_red_right", style);
		
		style = createBorderedStyle(wb);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setWrapText(false);
		style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldRedFont);
		styles.put("cell_highlight_right", style);

		return styles;
	}
	
	/**
	 * Helper method for the {@code createStyles} method above.  This was
	 * taken directly from example code included in the Apache POI project.
	 * Just creates a cell style that has a thin, black box border.
	 * 
	 * @param wb
	 * 		The workbook for which this style should be created.
	 * @return
	 */
	private static CellStyle createBorderedStyle(Workbook wb){
		BorderStyle thin = BorderStyle.THIN;
		short black = IndexedColors.BLACK.getIndex();

		CellStyle style = wb.createCellStyle();
		style.setBorderRight(thin);
		style.setRightBorderColor(black);
		style.setBorderBottom(thin);
		style.setBottomBorderColor(black);
		style.setBorderLeft(thin);
		style.setLeftBorderColor(black);
		style.setBorderTop(thin);
		style.setTopBorderColor(black);
		return style;
	}
}
