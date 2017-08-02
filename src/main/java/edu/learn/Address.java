package edu.learn;

import edu.learn.annotations.ExcelCell;

public class Address {

	@ExcelCell(headerName = "Address Line 1")
	private String addressLine1;
	@ExcelCell(headerName = "Zipcode")
	private int zipCode;

	public Address(String addressLine1, int zipCode) {
		this.addressLine1 = addressLine1;
		this.zipCode = zipCode;
	}

}
