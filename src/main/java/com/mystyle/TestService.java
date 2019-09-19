package com.mystyle;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestService {

	public static void main(String[] args) {
		// 1. Workbook creation
		// Workbook workbook = new SXSSFWorkbook(100);

		final String FILE_NAME = "SampleXLSFile_NewExcel.xlsx";
		// writeFile(FILE_NAME);
		System.out.println("Write task done!!!\n Now Reading...........");
		try {
			readExcelFile(FILE_NAME);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	private static void readExcelFile(String fileName) throws IOException {
		int i = 0;
		List<String[]> resultrowsList = new ArrayList<String[]>();
		FileInputStream excelInputStream = new FileInputStream(new File(fileName));
		Workbook workbook = new XSSFWorkbook(excelInputStream);
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowItr = sheet.iterator();
		int rowNum = 0;
		List<String> listeachrowdata = new ArrayList<String>();
		StringBuilder eachLine = new StringBuilder();
		while (rowItr.hasNext()) {

			Row row = rowItr.next();
			Iterator<Cell> cellItr = row.iterator();
			while (cellItr.hasNext()) {
				Cell cell = cellItr.next();

				if (cell.getCellType() == CellType.STRING) {
					// System.out.print(cell.getStringCellValue() + "\t");
					eachLine.append(cell.getStringCellValue());

				} else if (cell.getCellType() == CellType.NUMERIC) {
					// System.out.print(cell.getNumericCellValue() + "\t");
					eachLine.append(cell.getNumericCellValue());

				}
				if (cellItr.hasNext()) {
					eachLine.append(",");
				}

			}

			listeachrowdata.add(eachLine.toString());
			String[] rowline = listeachrowdata.toArray(new String[0]);
			eachLine.setLength(0);
			listeachrowdata.clear();
			resultrowsList.add(rowline);
			if (resultrowsList != null && (resultrowsList.size() > 100 && resultrowsList.size() > 0)) {
				i++;

				 printDataOfExcelFile(resultrowsList,i) ;
				resultrowsList.clear();

			}
			rowNum++;
		}
		workbook.close();
		excelInputStream.close();

	}

	private static void printDataOfExcelFile(List<String[]> rows, int k) throws IOException {

		System.out.println(
				"++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
						+ rows.size());
		for (int i = 0; i < rows.size(); i++) {
			String[] row = rows.get(i);
			System.out.println(row.toString());
			for (int j = 0; j < row.length; j++) {
				System.out.print(row[j]);
			}
			System.out.println();

		}
		rows.clear();
		System.out.println(
				"++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++end++++++++++"
						+ k);
	}

}
