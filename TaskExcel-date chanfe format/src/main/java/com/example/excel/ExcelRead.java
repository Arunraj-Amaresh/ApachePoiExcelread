package com.example.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelRead {

	// Method to read Excel (unchanged)
	public static List<HashMap<String, Object>> readExcel(String inputFilePath) {
		List<HashMap<String, Object>> dataList = new ArrayList<>();
		try (FileInputStream fileInputStream = new FileInputStream(new File(inputFilePath));
				Workbook workbook = new XSSFWorkbook(fileInputStream)) {

			Sheet sheet = workbook.getSheetAt(0);
			Row headerRow = sheet.getRow(0);
			List<String> headers = new ArrayList<>();

			for (Cell cell : headerRow) {
				headers.add(cell.getStringCellValue());
			}

			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				HashMap<String, Object> cellDataMap = new HashMap<>();
				for (int j = 0; j < headers.size(); j++) {
					Cell cell = row.getCell(j);
					switch (cell.getCellType()) {
					case STRING:
						cellDataMap.put(headers.get(j), cell.getStringCellValue());
						break;
					case NUMERIC:
						if (DateUtil.isCellDateFormatted(cell)) {
							cellDataMap.put(headers.get(j), cell.getDateCellValue());
						} else {
							cellDataMap.put(headers.get(j), cell.getNumericCellValue());
						}
						break;
					case BOOLEAN:
						cellDataMap.put(headers.get(j), cell.getBooleanCellValue());
						break;
					default:
						cellDataMap.put(headers.get(j), null);
					}
				}
				dataList.add(cellDataMap);
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
		return dataList;
	}

	public static void writeExcel(String outputFilePath, List<HashMap<String, Object>> dataList) {
		if (dataList.isEmpty()) {
			return; // No data to write
		}

		// Extract headers from the first row of data
		Set<String> headersSet = dataList.get(0).keySet();
		String[] headers = headersSet.toArray(new String[0]);

		try (Workbook outputWorkbook = new XSSFWorkbook();
				FileOutputStream fileOutputStream = new FileOutputStream(new File(outputFilePath))) {

			Sheet outputSheet = outputWorkbook.createSheet("Sheet1");
			Row headerRow = outputSheet.createRow(0);

			// Write header row
			for (int i = 0; i < headers.length; i++) {
				Cell headerCell = headerRow.createCell(i);
				headerCell.setCellValue(headers[i]);
			}

			int rowIdx = 1;
			for (HashMap<String, Object> rowDataMap : dataList) {
				Row outputRow = outputSheet.createRow(rowIdx++);

				for (int colIdx = 0; colIdx < headers.length; colIdx++) {
					Cell outputCell = outputRow.createCell(colIdx);
					Object value = rowDataMap.get(headers[colIdx]);

					if (value instanceof String) {
						outputCell.setCellValue((String) value);
					} else if (value instanceof Double) {
						if ("Salary".equals(headers[colIdx])) {
							outputCell.setCellValue(String.format("%.2f", (Double) value));
						} else {
							outputCell.setCellValue(((Double) value).intValue());
						}
					} else if (value instanceof Boolean) {
						outputCell.setCellValue((Boolean) value);
					} else if (value instanceof java.util.Date) {
						SimpleDateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy");
						outputCell.setCellValue(dateFormat.format((java.util.Date) value));
					} else if (value == null) {
						outputCell.setCellValue("");
					}
				}
			}

			outputWorkbook.write(fileOutputStream);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		String inputFilePath = "C:\\Users\\BSIT-021\\Documents\\workspace-spring-tool-suite-4-4.23.1.RELEASE\\TaskExcel\\src\\main\\resources\\sam.xlsx";
		String outputFilePath = "C:\\Users\\BSIT-021\\Documents\\workspace-spring-tool-suite-4-4.23.1.RELEASE\\TaskExcel\\src\\main\\resources\\output.xlsx";

		List<HashMap<String, Object>> dataList = readExcel(inputFilePath);
		writeExcel(outputFilePath, dataList);

		int rowNumber = 0;
		for (HashMap<String, Object> rowData : dataList) {
			System.out.println("Row " + rowNumber++ + ": " + rowData);
		}
	}
}
