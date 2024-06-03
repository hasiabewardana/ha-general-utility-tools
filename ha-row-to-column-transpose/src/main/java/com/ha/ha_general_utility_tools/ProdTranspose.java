package com.ha.ha_general_utility_tools;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ProdTranspose {
    public static void prodTranspose() {
        String inputFile = "D:\\Organizational\\Dialog_Axiata_PLC\\Work\\Projects\\MIFE\\MIFE_APIs-JBoss_Env-2024.05.30\\MIFE_APIs-Prod-JBoss_Env-v1.0.xlsx";  // Replace with your input file path
        String outputFile = "D:\\Organizational\\Dialog_Axiata_PLC\\Work\\Projects\\MIFE\\MIFE_APIs-JBoss_Env-2024.05.30\\MIFE_APIs-Prod-JBoss_Env-v2.0.xlsx"; // Replace with your desired output file path

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
            Cell outputCell = outputRow.createCell(0);
            outputCell.setCellValue("Owner");
            outputCell = outputRow.createCell(1);
            outputCell.setCellValue("API");
            outputCell = outputRow.createCell(2);
            outputCell.setCellValue("Prod Endpoint");
            outputCell = outputRow.createCell(3);
            outputCell.setCellValue("Prod Server");
            outputCell = outputRow.createCell(4);
            outputCell.setCellValue("Service");
            outputRow = outputSheet.createRow(++outputRowIndex);

            for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
                Row inputRow = inputSheet.getRow(rowIndex);

                if (inputRow != null && !inputRow.getCell(0).toString().equals("--")) {
                    Cell inputCell = inputRow.getCell(0);
                    outputCell = outputRow.createCell(colIndex++);

                    if (inputCell != null && colIndex == 1) {
                        if (inputCell.toString().startsWith("GITCRM--")) {
                            outputCell.setCellValue("GITCRM");
                        }
                    } else if (inputCell != null && colIndex == 2) {
                        if (inputCell.toString().startsWith("GITCRM--")) {
                            String prefix = "GITCRM--";
                            String suffix = ".xml-";

                            if (inputCell.toString().startsWith(prefix) && inputCell.toString().contains(suffix)) {
                                // Find the start index of the middle part
                                int startIndex = prefix.length();

                                // Find the end index of the middle part
                                int endIndex = inputCell.toString().indexOf(suffix);

                                // Extract and return the middle part
                                outputCell.setCellValue(inputCell.toString().substring(startIndex, endIndex).replace("_v", " - "));
                            }

                            String regex = "\"([^\"]*)\"";

                            // Compile the regex pattern
                            Pattern pattern = Pattern.compile(regex);

                            // Create a matcher for the input string
                            Matcher matcher = pattern.matcher(inputCell.toString());

                            String fullEndpoint = null;

                            // Find and return the matched substring (text within double quotes)
                            if (matcher.find()) {
                                outputCell = outputRow.createCell(colIndex++);
                                fullEndpoint = matcher.group(1);
                                outputCell.setCellValue(fullEndpoint); // group(1) returns the text within the quotes

                                // Define the regex pattern
                                String regex2 = "://([^/]+)/";

                                // Compile the regex pattern
                                Pattern pattern2 = Pattern.compile(regex2);

                                // Create a matcher for the input string
                                Matcher matcher2 = pattern2.matcher(fullEndpoint);

                                String server;

                                // Find and return the matched substring (text between delimiters)
                                if (matcher2.find()) {
                                    outputCell = outputRow.createCell(colIndex++);
                                    server = matcher2.group(1);
                                    outputCell.setCellValue(server); // group(1) returns the text within the capturing group

                                    server = server + "/";
                                    outputCell = outputRow.createCell(colIndex);
                                    int knownPartEndIndex = fullEndpoint.indexOf(server) + server.length();
                                    // Extract the substring from the end of the known part to the end of the string
                                    outputCell.setCellValue(fullEndpoint.substring(knownPartEndIndex));
                                } else {
                                    outputCell = outputRow.createCell(colIndex);
                                    outputCell.setCellValue(fullEndpoint.replace("http://", ""));
                                }
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
