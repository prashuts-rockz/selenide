package org.tests;

import org.apache.maven.model.Model;
import org.apache.maven.model.io.xpp3.MavenXpp3Reader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.*;
import java.lang.reflect.Method;
import java.net.URL;
import java.net.URLClassLoader;
import java.util.ArrayList;
import java.util.List;

public class HtmlExcelGenReport {

    public static void main(String[] args) {
        List<Class<?>> testClasses = new ArrayList<>();
        try {
            String projectPath = "target/test-classes"; // Adjust this path accordingly
            File rootDir = new File(projectPath);
            if (rootDir.exists() && rootDir.isDirectory()) {
                loadClasses(rootDir, projectPath, testClasses);
            } else {
                System.err.println("The specified path does not exist or is not a directory: " + projectPath);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Generate HTML report with test attributes
        generateHtmlReport(testClasses);
        generateExcelReport(testClasses);

    }

    private static void loadClasses(File directory, String rootPath, List<Class<?>> testClasses) throws IOException, ClassNotFoundException {
        URLClassLoader classLoader = URLClassLoader.newInstance(new URL[]{new File(rootPath).toURI().toURL()});

        if (directory.exists()) {
            File[] files = directory.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isDirectory()) {
                        loadClasses(file, rootPath, testClasses);
                    } else if (file.getName().endsWith(".class")) {
                        String className = getClassName(file, new File(rootPath));
                        Class<?> cls = Class.forName(className, true, classLoader);
                        testClasses.add(cls);
                    }
                }
            }
        }
    }

    private static String getClassName(File file, File rootPath) {
        String path = file.getAbsolutePath().replace(rootPath.getAbsolutePath(), "").replace(File.separator, ".");
        return path.startsWith(".") ? path.substring(1, path.length() - 6) : path.substring(0, path.length() - 6); // Remove leading period and ".class" extension
    }

    private static String getModuleName(Class<?> cls) {
        try {
            String classPath = cls.getProtectionDomain().getCodeSource().getLocation().getPath();
            File moduleDir = new File(classPath).getParentFile().getParentFile(); // Adjust as necessary
            File pomFile = new File(moduleDir, "pom.xml");

            if (pomFile.exists()) {
                MavenXpp3Reader reader = new MavenXpp3Reader();
                Model model = reader.read(new FileReader(pomFile));
                return model.getName();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "Unknown Module";
    }

    private static void generateHtmlReport(List<Class<?>> testClasses) {
        try (FileWriter writer = new FileWriter("test-attributes-report.html")) {
            writer.append("<html><body><h1>Test Attributes Report</h1>");
            writer.append("<table border='1'>");
            writer.append("<tr><th>Module Name</th><th>Test Class</th><th>Test Method</th><th>Description</th><th>Enabled</th><th>Groups</th><th>Priority</th><th>Timeout</th></tr>");

            for (Class<?> cls : testClasses) {
                String moduleName = getModuleName(cls);
                for (Method method : cls.getDeclaredMethods()) {
                    if (method.isAnnotationPresent(Test.class)) {
                        Test testAnnotation = method.getAnnotation(Test.class);

                        writer.append("<tr>");
                        writer.append("<td>").append(moduleName).append("</td>");
                        writer.append("<td>").append(cls.getName()).append("</td>");
                        writer.append("<td>").append(method.getName()).append("</td>");
                        writer.append("<td>").append(testAnnotation.description()).append("</td>");
                        writer.append("<td>").append(String.valueOf(testAnnotation.enabled())).append("</td>");
                        writer.append("<td>").append(String.join(", ", testAnnotation.groups())).append("</td>");
                        writer.append("<td>").append(String.valueOf(testAnnotation.priority())).append("</td>");
                        writer.append("<td>").append(String.valueOf(testAnnotation.timeOut())).append("</td>");
                        writer.append("</tr>");
                    }
                }
            }

            writer.append("</table>");
            writer.append("</body></html>");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void generateExcelReport(List<Class<?>> testClasses) {
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create a sheet and set up the title
            Sheet sheet = workbook.createSheet("Test Attributes Report");

            // Create title row
            Row titleRow = sheet.createRow(0);
            titleRow.createCell(0).setCellValue("Test Attributes Report");
            CellStyle titleStyle = workbook.createCellStyle();
            Font titleFont = workbook.createFont();
            titleFont.setBold(true);
            titleFont.setFontHeightInPoints((short) 16); // Larger font size
            titleStyle.setFont(titleFont);
            titleStyle.setAlignment(HorizontalAlignment.CENTER);
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            titleRow.getCell(0).setCellStyle(titleStyle);

            // Merge cells for the title
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7)); // Merging columns 0 to 7

            // Create header row with bold and colored background
            Row headerRow = sheet.createRow(1);
            String[] headers = {"Module Name", "Test Class", "Test Method", "Description", "Enabled", "Groups", "Priority", "Timeout"};
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 12); // Set font size for header
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // Set column widths to fit the content
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Start adding data rows
            int rowNum = 2; // Start filling from the 3rd row
            for (Class<?> cls : testClasses) {
                String moduleName = getModuleName(cls);
                for (Method method : cls.getDeclaredMethods()) {
                    if (method.isAnnotationPresent(Test.class)) {
                        Test testAnnotation = method.getAnnotation(Test.class);
                        Row row = sheet.createRow(rowNum++);
                        CellStyle dataStyle = workbook.createCellStyle();
                        dataStyle.setBorderBottom(BorderStyle.THIN);
                        dataStyle.setBorderTop(BorderStyle.THIN);
                        dataStyle.setBorderLeft(BorderStyle.THIN);
                        dataStyle.setBorderRight(BorderStyle.THIN);

                        // Fill in the data for each column
                        row.createCell(0).setCellValue(moduleName);
                        row.getCell(0).setCellStyle(dataStyle);

                        row.createCell(1).setCellValue(cls.getName());
                        row.getCell(1).setCellStyle(dataStyle);

                        row.createCell(2).setCellValue(method.getName());
                        row.getCell(2).setCellStyle(dataStyle);

                        row.createCell(3).setCellValue(testAnnotation.description());
                        row.getCell(3).setCellStyle(dataStyle);

                        row.createCell(4).setCellValue(testAnnotation.enabled());
                        row.getCell(4).setCellStyle(dataStyle);

                        row.createCell(5).setCellValue(String.join(", ", testAnnotation.groups()));
                        row.getCell(5).setCellStyle(dataStyle);

                        row.createCell(6).setCellValue(testAnnotation.priority());
                        row.getCell(6).setCellStyle(dataStyle);

                        row.createCell(7).setCellValue(testAnnotation.timeOut());
                        row.getCell(7).setCellStyle(dataStyle);
                    }
                }
            }

            // Write the Excel file
            try (FileOutputStream fileOut = new FileOutputStream("test-attributes-report.xlsx")) {
                workbook.write(fileOut);
            }

            System.out.println("Excel report generated successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void generateExcelReport1(List<Class<?>> testClasses) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test Attributes Report");
            Row headerRow = sheet.createRow(0);

            // Create header row
            String[] headers = {"Module Name", "Test Class", "Test Method", "Description", "Enabled", "Groups", "Priority", "Timeout"};
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            int rowNum = 1;

            for (Class<?> cls : testClasses) {
                String moduleName = getModuleName(cls);
                for (Method method : cls.getDeclaredMethods()) {
                    if (method.isAnnotationPresent(Test.class)) {
                        Test testAnnotation = method.getAnnotation(Test.class);
                        Row row = sheet.createRow(rowNum++);

                        row.createCell(0).setCellValue(moduleName);
                        row.createCell(1).setCellValue(cls.getName());
                        row.createCell(2).setCellValue(method.getName());
                        row.createCell(3).setCellValue(testAnnotation.description());
                        row.createCell(4).setCellValue(testAnnotation.enabled());
                        row.createCell(5).setCellValue(String.join(", ", testAnnotation.groups()));
                        row.createCell(6).setCellValue(testAnnotation.priority());
                        row.createCell(7).setCellValue(testAnnotation.timeOut());
                    }
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream("test-attributes-report.xlsx")) {
                workbook.write(fileOut);
            }

            System.out.println("Excel report generated successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
