package org.tests;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.Iterator;

public class ExcelToCSVConverter {

    public static void main(String[] args) {
        String excelFilePath = "excelpOC\\sample.xlsx"; // Path to Excel file
        String outputFolder = "excelpOC\\output_csv"; // Output directory for CSV files

        convertExcelToCSV(excelFilePath, outputFolder);
    }


        public static void convertExcelToCSV(String excelFilePath, String outputFolder) {
            try {
                File folder = new File(outputFolder);
                if (!folder.exists()) {
                    folder.mkdirs();
                }

                FileInputStream fileInputStream = new FileInputStream(excelFilePath);
                Workbook workbook = new XSSFWorkbook(fileInputStream);
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator(); // Formula Evaluator

                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();
                    File csvFile = new File(outputFolder + File.separator + sheetName + ".csv");

                    writeSheetToCSV(sheet, csvFile, evaluator);
                    System.out.println("Converted sheet: " + sheetName + " -> " + csvFile.getAbsolutePath());
                }

                workbook.close();
                fileInputStream.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        private static void writeSheetToCSV(Sheet sheet, File csvFile, FormulaEvaluator evaluator) {
            try (PrintWriter writer = new PrintWriter(new FileWriter(csvFile))) {
                for (Row row : sheet) {
                    StringBuilder sb = new StringBuilder();
                    for (Cell cell : row) {
                        String cellValue = getCellValueAsString(cell, evaluator);
                        sb.append(cellValue).append(",");
                    }

                    if (sb.length() > 0) {
                        sb.setLength(sb.length() - 1); // Remove last comma
                    }
                    writer.println(sb);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        private static String getCellValueAsString(Cell cell, FormulaEvaluator evaluator) {
            if (cell == null) return "";

            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString(); // Convert Date to String
                    }
                    return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return evaluateFormula(cell, evaluator); // Evaluate the formula
                case BLANK:
                default:
                    return "";
            }
        }

        private static String evaluateFormula(Cell cell, FormulaEvaluator evaluator) {
            try {
                CellValue cellValue = evaluator.evaluate(cell);
                switch (cellValue.getCellType()) {
                    case STRING:
                        return cellValue.getStringValue();
                    case NUMERIC:
                        return String.valueOf(cellValue.getNumberValue());
                    case BOOLEAN:
                        return String.valueOf(cellValue.getBooleanValue());
                    default:
                        return "";
                }
            } catch (Exception e) {
                return "ERROR"; // If formula evaluation fails
            }
        }
    }
