package com.example.excel;




import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class ExcelRead {

    public static HashMap<Integer, List<Object>> readAndWriteExcel(String inputFilePath, String outputFilePath) {
        HashMap<Integer, List<Object>> dataMap = new HashMap<>();

        try (FileInputStream fileInputStream = new FileInputStream(new File(inputFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            int rowNumber = 0;

            for (Row row : sheet) {
                List<Object> cellData = new ArrayList<>();
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            cellData.add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            cellData.add(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            cellData.add(cell.getBooleanCellValue());
                            break;
                        default:
                            cellData.add(null);
                    }
                }
                dataMap.put(rowNumber++, cellData);
            }

            // Now write the data to a new Excel file
            try (Workbook outputWorkbook = new XSSFWorkbook();
                 FileOutputStream fileOutputStream = new FileOutputStream(new File(outputFilePath))) {

                Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

                for (int rowIdx : dataMap.keySet()) {
                    Row outputRow = outputSheet.createRow(rowIdx);
                    List<Object> rowData = dataMap.get(rowIdx);

                    for (int colIdx = 0; colIdx < rowData.size(); colIdx++) {
                        Cell outputCell = outputRow.createCell(colIdx);
                        Object value = rowData.get(colIdx);

                        if (value instanceof String) {
                            outputCell.setCellValue((String) value);
                        } else if (value instanceof Double) {
                            outputCell.setCellValue((Double) value);
                        } else if (value instanceof Boolean) {
                            outputCell.setCellValue((Boolean) value);
                        }
                    }
                }

                outputWorkbook.write(fileOutputStream);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return dataMap;
    }

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\BSIT-021\\Documents\\workspace-spring-tool-suite-4-4.23.1.RELEASE\\TaskExcel\\src\\main\\resources\\sam.xlsx";
        String outputFilePath = "C:\\Users\\BSIT-021\\Documents\\workspace-spring-tool-suite-4-4.23.1.RELEASE\\TaskExcel\\src\\main\\resources\\output.xlsx";

        HashMap<Integer, List<Object>> dataMap = readAndWriteExcel(inputFilePath, outputFilePath);

        // Print the data from HashMap
        dataMap.forEach((key, value) -> {
            System.out.println("Row " + key + ": " + value);
        });
    }
}
