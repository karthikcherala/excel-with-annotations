package edu.learn;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {

	public static void main(String[] args) {
		try {
			System.out.println("Executing...");
			List<Student> students = getTestData();
			ExcelWriter excelWriter = new ExcelWriter("/Users/khome/ExcelFile.xlsx");
			excelWriter.writeSheet(students, Student.class);
			System.out.println("Success");
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	static List<Student> getTestData() {
		List<Student> students = new ArrayList<Student>();
		students.add(new Student(1, "Barry", "Allen", "Central City", new Address("A2", 2)));
		students.add(new Student(2, "Oliver", "Queen", "Starling City", new Address("A3", 3)));
		students.add(new Student(3, "Clark", "Kent", "Metropolis", new Address("A1", 1)));
		students.add(new Student(4, "Peter", "Parker", "New York City", new Address("A4", 4)));
		return students;
	}
}
