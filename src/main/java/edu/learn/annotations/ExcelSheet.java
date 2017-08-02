package edu.learn.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(value = RetentionPolicy.RUNTIME)
public @interface ExcelSheet {
	
	/**
	 * Name of the Excel sheet 
	 */
	public String value();
}
