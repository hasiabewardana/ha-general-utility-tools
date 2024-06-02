package com.ha.ha_general_utility_tools;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        System.out.print("Hello and welcome to HA Row to Column Transpose!\n");

        String inputFile = "D:\\Organizational\\Dialog_Axiata_PLC\\Work\\Projects\\MIFE\\MIFE_APIs-JBoss_Env-2024.05.30\\MIFE_APIs-JBoss_Env-v1.0.xlsx";  // Replace with your input file path
        String outputFile = "D:\\Organizational\\Dialog_Axiata_PLC\\Work\\Projects\\MIFE\\MIFE_APIs-JBoss_Env-2024.05.30\\MIFE_APIs-JBoss_Env-v2.0.xlsx"; // Replace with your desired output file path

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(fis);
             FileOutputStream fos = new FileOutputStream(outputFile);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet inputSheet = workbook.getSheetAt(0); // Assuming data is in the first sheet
            Sheet outputSheet = outputWorkbook.createSheet(); // Create a new sheet in the output workbook

            int numRows = inputSheet.getLastRowNum() + 1;

            System.out.printf("Number of Rowa: %d\n", numRows);

            int outputRowIndex = 0;
            int colIndex = 0;
            Row outputRow = outputSheet.createRow(outputRowIndex);

            for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
                Row inputRow = inputSheet.getRow(rowIndex);

                if (inputRow != null) {
                    Cell inputCell = inputRow.getCell(0);
                    Cell outputCell = outputRow.createCell(colIndex++);
                    if (inputCell != null) {
                        outputCell.setCellValue(inputCell.toString());
                    } else {
                        outputCell.setCellValue("");
                    }
                } else {
                    colIndex = 0;
                    outputRow = outputSheet.createRow(++outputRowIndex);
                }
            }

            outputWorkbook.write(fos); // Write the output workbook to the output file

            System.out.println("Excel transpose complete. Output written to " + outputFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}