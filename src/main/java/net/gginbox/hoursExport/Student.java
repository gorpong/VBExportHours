package net.gginbox.hoursExport;

import java.util.HashMap;
import java.util.regex.Pattern;

/**
 * Class that contains the people-specific aspects of the exported hours listing.
 * This is a data structure for a person in the Excel export file.  The concept
 * of the team that they are on is a higher-order function to be handled by
 * either another class that will include Student as one of its attributes or
 * by just manually creating something like a HashMap that has Student as one of
 * its elements.  This implements the Comparable interface so that we can use
 * the lambda functions to quickly sort the people based on the hours they
 * worked.
 * 
 * @author Gordon Galligher - gorpong@gginbox.net
 *
 */
public class Student implements Comparable<Student> {
	
	/*
	 * Class variables
	 */
	
	private static HashMap<String, Student> _students = new HashMap<String, Student>();
	
	/*
	 * Instance variables
	 */
	private String lname;			// Last name
	private String fname;			// First name
	private String id;				// ID number (e.g., their school ID number)
	private double hours = 0.0;		// The number of hours worked
	
		
	/**
	 * Private constructor, use the getStudent() factory method instead.
	 * 
	 * @param name
	 * 		The student's full name (or whatever is in the file)
	 * @param id
	 * 		The student's ID (this is the only value on parent's report)
	 * @param hours
	 * 		The hours for that week.
	 */
/*
	private Student(String name, String id, double hours) {
		Pattern regexComma = Pattern.compile(".*, .*");
		Pattern regexSpace = Pattern.compile(".* .*");
		this.id    = id;
		this.hours = hours;
		if ( regexComma.matcher(name).matches() ) {
			this.lname = name.substring(0, name.indexOf(','));
			this.fname = name.substring(name.indexOf(',')+2);
		} else if ( regexSpace.matcher(name).matches() ) {
			this.lname = name.substring(0, name.indexOf(' '));
			this.fname = name.substring(name.indexOf(' ')+2);
		} else {
			this.lname = name;
			this.fname = "";
		}
	}
*/
	private Student(String name, String id, double hours) {
	    Pattern regexComma = Pattern.compile(".*, .*");
	    Pattern regexSpace = Pattern.compile(".* .*");
	    this.id = id;
	    this.hours = hours;
	    if (regexComma.matcher(name).matches()) {
	        String[] parts = name.split(", ");
	        this.lname = parts[0].trim();
	        this.fname = parts[1].trim();
	    } else if (regexSpace.matcher(name).matches()) {
	        String[] parts = name.split(" ");
	        this.lname = parts[1].trim();
	        this.fname = parts[0].trim();
	    } else {
	        this.lname = name;
	        this.fname = "";
	    }
	}
	
	/**
	 * Factory method to get a Student object, if one for that id exists, then add
	 * the hours to its existing hours, otherwise create a new one, and then return it.
	 * 
	 * @param name
	 * 		The name of the student
	 * @param id
	 * 		The ID of the student (must be unique across all Student instances)
	 * @param hours
	 * 		The hours for that particular time period
	 * @return
	 * 		The previously created/newly created Student object for that person
	 */
	public static Student getStudent(String name, String id, double hours) {
		Student person;
		if ( _students.containsKey(id) ) {
			person = _students.get(id);
			person.hours += hours;
		} else {
			person = new Student(name, id, hours);
			_students.put(id,  person);
		}
		return person;
	}

	public String getName() {
		return lname + ", " + fname;
	}
	public String getId() {
		return id;
	}
	public double getHours() {
		return hours;
	}
	
	/**
	 * Implement the Comparator method so we can compare two People and
	 * figure out which one is "more" based on the hours they've worked.
	 * This is a descending sort in that at the end of your list, you'll
	 * have them sorted in most hours to least hours.
	 */
	public int compareTo (Student that) {
		if ( this.hours > that.hours ) return -1;
		else if ( this.hours == that.hours ) return 0;
		else return 1;
	}

	public static void clearStudents() {
		_students.clear();
	}
	/**
	 * Pretty-print the structure in the form:  [firstname,lastname:id:hours]
	 */
	public String toString() {
		return "[" + fname + " " + lname + ":" + id + ":" + hours + "]";
	}
}