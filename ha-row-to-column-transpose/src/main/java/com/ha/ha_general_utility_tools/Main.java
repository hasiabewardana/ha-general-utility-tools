package com.ha.ha_general_utility_tools;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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

            int outputRowIndex = 1;
            int colIndex = 0;
            Row outputRow = outputSheet.createRow(outputRowIndex);

            for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
                Row inputRow = inputSheet.getRow(rowIndex);

                if (inputRow != null) {
                    Cell inputCell = inputRow.getCell(0);
                    Cell outputCell = outputRow.createCell(colIndex++);

                    if (inputCell != null && colIndex == 1) {
                        if (inputCell.toString().startsWith("Context: ")) {
                            outputCell.setCellValue(inputCell.toString().substring("Context: ".length()));
                        }
                    } else if (inputCell != null && colIndex == 2) {
                        Row thirdInputRow = inputSheet.getRow(rowIndex + 1);
                        Cell inputCellThirdInputRow;

                        if (thirdInputRow != null) {
                            inputCellThirdInputRow = thirdInputRow.getCell(0);

                            if (Objects.equals(inputCell.toString(), "Endpoints and URI Templates:")
                                    && inputCellThirdInputRow.toString().startsWith("GITCRM")) {
                                outputCell.setCellValue("GITCRM");
                            }
                        }
                    } else if (inputCell != null && colIndex == 3) {
                        if (inputCell.toString().startsWith("GITCRM--")) {
                            String temp1 = inputCell.toString().substring("GITCRM--".length());
                            String temp2;

                            if (temp1.contains("_APIproductionEndpoint")) {
                                int position = temp1.indexOf("_APIproductionEndpoint");

                                if (position != -1) {
                                    temp2 = temp1.substring(0, position);
                                    outputCell.setCellValue(temp2);
                                }

                                String regex = "\\s+(\\S+)";
                                Pattern pattern = Pattern.compile(regex);
                                Matcher matcher = pattern.matcher(temp1);

                                if (matcher.find()) {
                                    outputCell = outputRow.createCell(colIndex++);
                                    outputCell.setCellValue(matcher.group(1)); // Return the matched text
                                }
                            } else if (temp1.contains("_APIsandboxEndpoint")) {
                                int position = temp1.indexOf("_APIsandboxEndpoint");

                                if (position != -1) {
                                    temp2 = temp1.substring(0, position);
                                    outputCell.setCellValue(temp2);
                                }

                                String regex = "\\s+(\\S+)";
                                Pattern pattern = Pattern.compile(regex);
                                Matcher matcher = pattern.matcher(temp1);

                                if (matcher.find()) {
                                    outputCell = outputRow.createCell(colIndex + 1);
                                    outputCell.setCellValue(matcher.group(1)); // Return the matched text
                                }
                            }
                        }
                    } else if (inputCell != null && colIndex == 5) {
                        if (inputCell.toString().contains("_APIsandboxEndpoint")) {
                            String regex = "\\s+(\\S+)";
                            Pattern pattern = Pattern.compile(regex);
                            Matcher matcher = pattern.matcher(inputCell.toString());

                            if (matcher.find()) {
                                outputCell.setCellValue(matcher.group(1)); // Return the matched text
                            }
                        }
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