package com.ha.ha_general_utility_tools;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

public class RemoveDuplicateRows {
    public static void removeDuplicateRows() {
        String inputFilePath = "D:\\Organizational\\Dialog_Axiata_PLC\\Work\\Projects\\MIFE\\MIFE_APIs-JBoss_Env-2024.05.30\\MIFE_APIs-Prod-JBoss_Env-v2.0.xlsx";
        String outputFilePath = "D:\\Organizational\\Dialog_Axiata_PLC\\Work\\Projects\\MIFE\\MIFE_APIs-JBoss_Env-2024.05.30\\MIFE_APIs-Prod-JBoss_Env-v2.0.xlsx";

        try {
            // Load the Excel file
            FileInputStream fileInputStream = new FileInputStream(inputFilePath);
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);

            // Use a Set to track unique rows
            Set<String> uniqueRows = new HashSet<>();
            int lastRowNum = sheet.getLastRowNum();

            // Iterate through the rows and remove duplicates
            for (int i = 0; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Convert the row to a string representation
                StringBuilder rowString = new StringBuilder();
                for (Cell cell : row) {
                    rowString.append(cell.toString()).append(",");
                }

                // Check if the row is unique
                if (!uniqueRows.add(rowString.toString())) {
                    // If the row is a duplicate, remove it
                    removeRow(sheet, i);
                    i--; // Adjust the index after removing the row
                    lastRowNum--; // Adjust the last row number after removing the row
                }
            }

            // Write the updated data back to a new Excel file
            FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
            workbook.write(fileOutputStream);
            workbook.close();
            fileInputStream.close();
            fileOutputStream.close();

            System.out.println("Duplicate rows removed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }
}