package edu.learn.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(value = RetentionPolicy.RUNTIME)
public @interface ExcelCell {
	
	public int columnOrder() default 0;
	public String headerName();

}
