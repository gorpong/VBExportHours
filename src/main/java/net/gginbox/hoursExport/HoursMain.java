package net.gginbox.hoursExport;

import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.commons.cli.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import net.gginbox.hoursExport.Output.SheetType;
import java.util.Calendar;

/**
 * Program to scan an Excel-formatted file from a fingerprint scanner and
 * organize the data by Team ID, Student ID and number hours worked that week.
 * This is to for the Vandegrift ViperBots FTC Robotics program at Vandegrift
 * High School in Austin, Texas.
 *  
 * @author Gordon Galligher - gorpong@gginbox.net
 *
 */
public class HoursMain {
	
	/**
	 * Main program, read Excel file hours export file, organize by team.
	 * 
	 * @param argv<br>
	 * 		-c cfgFile -- The configuration file to read (default:  VBHoursExport.propties)<br>
	 * 		-i infile  -- The input file to read containing source hours<br>
	 * 		-o outfile -- The output file to write/create, containing formatted data<br>
	 * 		-lt hrs	   -- The low hours threshold, under which are flagged in bold-red<br>
	 * 	    -ht hrs	   -- The high hours threshold, over which are highlighted, bold-red<br>
	 * <p>
	 * All command line arguments override any configuration file settings for similar values.
	 * </p>
	 * 
	 * @throws IOException 			 Error when closing workbook
	 * @throws FileNotFoundException Can't open file or can't write file
	 */
	public static void main(String[] argv) throws FileNotFoundException, IOException {
		Options options = new Options();
		options.addOption( Option.builder("c").hasArg()
				.argName("configuration file")
				.longOpt("cfgFile")
				.desc("Configuration file path")
				.build());
		options.addOption( Option.builder("i").hasArg()
				.argName("input file")
				.longOpt("inFile")
				.desc("Input file")
				.build());
		options.addOption( Option.builder("o").hasArg()
				.argName("output file")
				.longOpt("outFile")
				.desc("Output file")
				.build());
		options.addOption( Option.builder("lt").hasArg()
				.argName("low threshold")
				.longOpt("lowThreshold")
				.desc("Threshold for hours too low")
				.build());
		options.addOption( Option.builder("ht").hasArg()
				.argName("high threshold")
				.longOpt("highThreshold")
				.desc("Threshold for hours too high")
				.build());
		CommandLineParser parser = new DefaultParser();
		CommandLine cmd = null;
		try {
			cmd = parser.parse(options, argv);
		}
		catch ( org.apache.commons.cli.ParseException exp ) {
			System.out.println(exp.getMessage());
			HelpFormatter formatter = new HelpFormatter();
			formatter.printHelp("VBHoursExport", options);
			System.exit(1);
		}
		
		ConfigProperties config = new ConfigProperties();
		if ( cmd.hasOption("c") )
			config.getPropValues(cmd.getOptionValue('c'));
		else
			config.getPropValues();
		
		Calendar date = Calendar.getInstance();
		String datestr = String.format("%02d/%02d/%04d %02d:%02d %s",  date.get(Calendar.MONTH)+1, 
				date.get(Calendar.DATE), date.get(Calendar.YEAR), 
				date.get(Calendar.HOUR), date.get(Calendar.MINUTE), 
				date.get(Calendar.AM_PM) == 1 ? "PM" : "AM");
		System.out.println("Processing starting at:  " + datestr);
		String inputFile  = cmd.getOptionValue("i", config.getConfig("inputFile"));
		System.out.println("Reading from file:  " + inputFile);
		Teams scanner = new Teams(inputFile, config);
		try {
			scanner.parseExcel();
			// scanner.printTeamInfo();
		} catch (EncryptedDocumentException e) {
			System.err.println("Cannot parse encrypted documents");
			e.printStackTrace();
			System.exit(1);
		} catch (InvalidFormatException e) {
			System.out.println("Not all columns present in " + inputFile + ": " + e.getMessage());
			System.exit(1);
			e.printStackTrace();
		}
		int numStudents = 0;
		int numTeams = 0;
		for (Integer team : scanner.getTeams()) {
			numTeams++;
			numStudents += scanner.getHoursByTeam(team).size();
		}
		System.out.print(String.format("Successfully parsed Input file:  %d teams and %d students\n", numTeams, numStudents)); 
				
		String outputFile = cmd.getOptionValue("o", config.getConfig("outputFile")); 
		double hrsLow = 0.0;  
		double hrsHigh = 0.0;
		try {
			hrsLow  = Double.parseDouble(cmd.getOptionValue("lt", config.getConfig("hoursLowThreshold", "3.0")));
			hrsHigh = Double.parseDouble(cmd.getOptionValue("ht", config.getConfig("hoursHighThreshold", "7.0")));
		} catch (NumberFormatException e) {
			System.err.println("Illegal number format for Config hoursLow/HighThreshold and/or -lt/-ht command args");
			System.exit(1);
		}
		Output out = Output.initialize(outputFile, config, hrsLow, hrsHigh);
		out.createSheet(scanner, SheetType.COACHES);
		out.createSheet(scanner, SheetType.PARENTS);
		out.close();
		System.out.print(String.format("Noted %d students with low hours and %d students with high hours\n",
				out.getLowCount(), out.getHighCount()));
		date = Calendar.getInstance();
		datestr = String.format("%02d/%02d/%04d %02d:%02d %s",  date.get(Calendar.MONTH)+1, 
				date.get(Calendar.DATE), date.get(Calendar.YEAR), 
				date.get(Calendar.HOUR), date.get(Calendar.MINUTE), 
				date.get(Calendar.AM_PM) == 1 ? "PM" : "AM");
		System.out.println("Processing Complete at:  " + datestr);
		System.out.println("File Created:  " + outputFile);
	}
	
}
