package edu.learn;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import edu.learn.annotations.ExcelCell;
import edu.learn.annotations.ExcelSheet;

public class ExcelWriter {

	private Workbook workbook;
	private File file;

	public ExcelWriter(String fileName) throws IOException {
		workbook = new XSSFWorkbook();
		file = new File(fileName);
		if (!file.exists()) {
			file.createNewFile();
		}
	}

	public void writeSheet(List<? extends Object> data, Class t) {
		if (data == null || data.isEmpty()) {
			System.out.println("Empty Data");
			return;
		}

		// I defined another annotation in the project that can be applied to
		// Student class to specify the name of the sheet represented by the
		// Student class.
		ExcelSheet sheetAnnotation = (ExcelSheet) t.getAnnotation(ExcelSheet.class);
		Sheet sheet;
		if (sheetAnnotation == null)
			sheet = workbook.createSheet();
		else
			sheet = workbook.createSheet(sheetAnnotation.value());

		// Retreive all the properties of class which are annotated with
		// ExcelCell. Except the property 'extraData' all other properties of
		// the class will be collected here
		List<Field> excelColumns = getAnnonatedColumns(t);
		int rowCount = 0;
		int colCount = 0;

		// Write Headers
		Row headerRow = sheet.createRow(rowCount++);
		colCount = writeHeaders(colCount, excelColumns, "", headerRow);

		// Write Data
		for (Object dataObject : data) {
			Row dataRow = sheet.createRow(rowCount++);
			colCount = 0;
			colCount = writeRowData(colCount, dataRow, excelColumns, dataObject);
		}

		// Autosizing columns to cleanup the appearance of the spreadsheet.
		for (int i = 0; i < colCount; i++) {
			sheet.autoSizeColumn(i);
		}

		FileOutputStream fileOutputStream = null;

		try {
			fileOutputStream = new FileOutputStream(file);
			workbook.write(fileOutputStream);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (fileOutputStream != null) {
				try {
					fileOutputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	CellStyle getHeaderStyle() {
		CellStyle headerStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 12);
		font.setBold(true);
		font.setColor(IndexedColors.WHITE.getIndex());

		headerStyle.setFont(font);
		headerStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		headerStyle.setBorderBottom((short) 1);
		headerStyle.setBorderLeft((short) 1);
		headerStyle.setBorderRight((short) 1);
		headerStyle.setBorderTop((short) 1);

		return headerStyle;
	}

	CellStyle getDataStyle() {
		CellStyle dataStyle = workbook.createCellStyle();
		dataStyle.setBorderBottom((short) 1);
		dataStyle.setBorderLeft((short) 1);
		dataStyle.setBorderRight((short) 1);
		dataStyle.setBorderTop((short) 1);
		return dataStyle;
	}

	int writeHeaders(int startCol, List<Field> excelColumns, String parent, Row headerRow) {
		CellStyle headerStyle = getHeaderStyle();
		for (Field header : excelColumns) {

			ExcelCell cellAnnotation = header.getAnnotation(ExcelCell.class);
			String headerLabel = parent + cellAnnotation.headerName();
			if (header.getType().isPrimitive() || header.getType().equals(String.class)) {
				Cell cell = headerRow.createCell(startCol++);
				cell.setCellValue(headerLabel);
				cell.setCellStyle(headerStyle);
			} else {
				List<Field> subColumns = getAnnonatedColumns(header.getType());
				startCol = writeHeaders(startCol, subColumns, headerLabel + ".", headerRow);
			}
		}
		return startCol;
	}

	int writeRowData(int colCount, Row dataRow, List<Field> excelColumns, Object data) {
		CellStyle dataStyle = getDataStyle();

		for (Field field : excelColumns) {
			try {
				if (field.getType().isPrimitive() || field.getType().equals(String.class)) {

					String value = String.valueOf(field.get(data));
					Cell cell = dataRow.createCell(colCount++);
					cell.setCellValue(value);
					cell.setCellStyle(dataStyle);
				} else {
					List<Field> subColumns = getAnnonatedColumns(field.getType());
					colCount = writeRowData(colCount, dataRow, subColumns, field.get(data));
				}
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		return colCount;
	}

	List<Field> getAnnonatedColumns(Class t) {
		List<Field> excelColumns = new ArrayList<>();
		for (Field field : t.getDeclaredFields()) {
			ExcelCell cellAnnotation = field.getAnnotation(ExcelCell.class);
			if (cellAnnotation != null) {
				field.setAccessible(true);
				excelColumns.add(field);
			}
		}
		return excelColumns;
	}

}
