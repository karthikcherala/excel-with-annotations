package edu.learn;

import edu.learn.annotations.ExcelCell;

public class Student {

	@ExcelCell(headerName = "Roll Number", columnOrder = 1)
	private int rollNumber;
	@ExcelCell(headerName = "First Name", columnOrder = 2)
	private String firstName;
	@ExcelCell(headerName = "Last Name", columnOrder = 3)
	private String lastName;
	@ExcelCell(headerName = "City", columnOrder = 4)
	private String city;
	
	@ExcelCell(headerName = "Address", columnOrder = 5)
	private Address address;

	// Just to demonstrate that this class can also contain columns that we do
	// not want to be a part of excel
	private String extraData;

	public Student(int rollNumber, String firstName, String lastName, String city, Address address) {
		this.rollNumber = rollNumber;
		this.firstName = firstName;
		this.lastName = lastName;
		this.city = city;
		this.extraData = "Any other data";
		this.address = address;
	}
	
	// Getters and Setters for all members

	public int getRollNumber() {
		return rollNumber;
	}

	public void setRollNumber(int rollNumber) {
		this.rollNumber = rollNumber;
	}

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public String getLastName() {
		return lastName;
	}

	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	public String getCity() {
		return city;
	}

	public void setCity(String city) {
		this.city = city;
	}
	
	

}
