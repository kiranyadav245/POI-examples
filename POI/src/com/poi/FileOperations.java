package com.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class FileOperations {
	private static Workbook getWorkbook(String excelFilePath) throws IOException {
		Workbook workbook = null;
		if (excelFilePath.endsWith("xls")) {
			workbook = new HSSFWorkbook();
		} else {
			throw new IllegalArgumentException("The specified file is not Excel file");
		}
		return workbook;
	}

	public static void main(String[] args) throws IOException {
		Workbook workbook = getWorkbook("xls");
		Sheet sheet = workbook.createSheet("DB Data");

		List<DataVo> list = new ArrayList<DataVo>();
		for (int i = 0; i <= 10; i++) {
			DataVo d = new DataVo("name" + i, "id" + 1);
			list.add(d);
		}
		List<String> headings = new ArrayList<String>();
		headings.add("Name");
		headings.add("Id");
		int rowCount = 0;
		CellStyle cellHeaderStyle = sheet.getWorkbook().createCellStyle();
		Font hFont = sheet.getWorkbook().createFont();
	    hFont.setFontHeightInPoints((short) 16);
	    cellHeaderStyle.setFont(hFont);
	    cellHeaderStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
	    cellHeaderStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
	    cellHeaderStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
	    cellHeaderStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
	    cellHeaderStyle.setFillPattern(HSSFCellStyle.BORDER_THICK);
	    cellHeaderStyle.setFillBackgroundColor(HSSFColor.RED.index);
	    cellHeaderStyle.setFillForegroundColor(HSSFColor.RED.index);
	    cellHeaderStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	    Row row = sheet.createRow(rowCount++);
	    int Count = 0;
		for (String heading : headings) {
			Cell cell = row.createCell(Count++);
			cell.setCellValue(heading);
			cell.setCellStyle(cellHeaderStyle);
		}
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		Font font = sheet.getWorkbook().createFont();
	    font.setFontHeightInPoints((short) 10);
	    cellStyle.setFont(font);
	    cellStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
	    cellStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
	    cellStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
	    cellStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
	    cellStyle.setFillPattern(HSSFCellStyle.BORDER_THICK);
		for (DataVo vo : list) {
			row = sheet.createRow(rowCount++);
			int columnCount = 0;
			Cell cell = row.createCell(columnCount++);
			cell.setCellValue(vo.getName());
			cell.setCellStyle(cellStyle);
			Cell cell2 = row.createCell(columnCount++);
			cell2.setCellValue(vo.getId());
			cell2.setCellStyle(cellStyle);
		}

		try (FileOutputStream outputStream = new FileOutputStream("e:\\samples.xls")) {
			workbook.write(outputStream);
		}
		
	}
}

class DataVo {
	private String name;
	private String id;

	public String getName() {
		return name;
	}

	public DataVo() {

	}

	public DataVo(String name, String id) {
		this.name = name;
		this.id = id;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}
}