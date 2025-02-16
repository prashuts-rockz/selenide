package org.tests;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.annotations.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
public class ExcelReportTest {
    private static final String FILE_PATH = "src/test/java/resources/Sample.xlsx";
    private Workbook workbook;
    private Sheet sheet;

    @BeforeClass
    public void setUp() throws IOException {
        FileInputStream file = new FileInputStream(new File(FILE_PATH));
        workbook = new XSSFWorkbook(file);
        sheet = workbook.getSheet("Mid Month Analysis");
    }


    @Test
    public void testSummaryInfoSection() {
        Row row = sheet.getRow(0);
        Assert.assertEquals(row.getCell(0).getStringCellValue(), "Summary info", "Incorrect section name");
        Assert.assertEquals(sheet.getRow(1).getCell(0).getStringCellValue(), "Account1", "Missing 'Account1' section");
    }

    @Test
    public void testAccountWiseDetails() {
        Assert.assertEquals(sheet.getRow(2).getCell(0).getStringCellValue(), "Account wise details:", "Missing 'Account wise details' section");
        Assert.assertEquals(sheet.getRow(3).getCell(0).getStringCellValue(), "Name of the account holder", "Incorrect field name");
        Assert.assertEquals(sheet.getRow(4).getCell(0).getStringCellValue(), "Address", "Incorrect field name");
        Assert.assertEquals(sheet.getRow(5).getCell(0).getStringCellValue(), "Name of the Bank", "Incorrect field name");
        Assert.assertEquals(sheet.getRow(6).getCell(0).getStringCellValue(), "Account number 1", "Incorrect field name");
    }

    @Test
    public void testMidMonthAnalysisSheetExists() {
        Sheet midMonthSheet = workbook.getSheet("Mid Month Analysis");
        Assert.assertNotNull(midMonthSheet, "Sheet 'Mid Month Analysis' is missing in the Excel file");
    }

    @AfterClass
    public void tearDown() throws IOException {
        workbook.close();
    }

/*

    @Test
    public void testDefaultDateRange() {
        boolean monthwiseSectionFound = false;
        String[] expectedDateRanges = {
                "02/14/2025-03/13/2025",
                "03/14/2025-04/13/2025",
                "04/14/2025-05/13/2025",
                "05/14/2025-06/13/2025",
                "06/14/2025-07/13/2025",
                "07/14/2025-08/13/2025"
        };
        String[] actualDateRanges = new String[6];

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell firstCell = row.getCell(0);
            if (firstCell != null) {
                String cellValue = firstCell.getStringCellValue().replaceAll("\\s+", " ").trim();
                System.out.println("Checking row " + i + ": " + cellValue);

                if ("Monthwise Details".equalsIgnoreCase(cellValue)) {
                    monthwiseSectionFound = true;
                    Row nextRow = sheet.getRow(i + 1);
                    if (nextRow != null) {
                        for (int j = 0; j < 6; j++) {
                            Cell dateCell = nextRow.getCell(j + 1);
                            if (dateCell != null) {
                                if (dateCell.getCellType() == CellType.STRING) {
                                    actualDateRanges[j] = dateCell.getStringCellValue().trim();
                                } else if (dateCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateCell)) {
                                    actualDateRanges[j] = new DataFormatter().formatCellValue(dateCell);
                                }
                            }
                        }
                    }
                    break;
                }
            }
        }

        Assert.assertTrue(monthwiseSectionFound, "'Monthwise Details' section is missing in the Excel file");

        for (int k = 0; k < 6; k++) {
            System.out.println("Expected Date Range: " + expectedDateRanges[k]);
            System.out.println("Actual Date Range: " + actualDateRanges[k]);
            Assert.assertEquals(actualDateRanges[k], expectedDateRanges[k], "Date range is incorrect for column " + (k + 1));
        }
    }
*/

    @Test(description = "Verifies the default date range for six months and additional columns Total Average and Tag")
    public void testDefaultDateRange() {
        boolean monthwiseSectionFound = false;
        String[] expectedDateRanges = {
                "02/14/2025-03/13/2025",
                "03/14/2025-04/13/2025",
                "04/14/2025-05/13/2025",
                "05/14/2025-06/13/2025",
                "06/14/2025-07/13/2025",
                "07/14/2025-08/13/2025"
        };
        String[] actualDateRanges = new String[6];
        String actualTotalAverage = "";
        String actualTag = "";

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell firstCell = row.getCell(0);
            if (firstCell != null) {
                String cellValue = firstCell.getStringCellValue().replaceAll("\\s+", " ").trim();
                System.out.println("Checking row " + i + ": " + cellValue);

                if ("Monthwise Details".equalsIgnoreCase(cellValue)) {
                    monthwiseSectionFound = true;
                    Row nextRow = sheet.getRow(i + 1);
                    if (nextRow != null) {
                        for (int j = 0; j < 6; j++) {
                            Cell dateCell = nextRow.getCell(j + 1);
                            if (dateCell != null) {
                                if (dateCell.getCellType() == CellType.STRING) {
                                    actualDateRanges[j] = dateCell.getStringCellValue().trim();
                                } else if (dateCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateCell)) {
                                    actualDateRanges[j] = new DataFormatter().formatCellValue(dateCell);
                                }
                            }
                        }

                        Cell totalAverageCell = nextRow.getCell(7);
                        if (totalAverageCell != null) {
                            actualTotalAverage = totalAverageCell.getStringCellValue().trim();
                        }

                        Cell tagCell = nextRow.getCell(8);
                        if (tagCell != null) {
                            actualTag = tagCell.getStringCellValue().trim();
                        }
                    }
                    break;
                }
            }
        }

        Assert.assertTrue(monthwiseSectionFound, "'Monthwise Details' section is missing in the Excel file");

        for (int k = 0; k < 6; k++) {
            System.out.println("Expected Date Range: " + expectedDateRanges[k]);
            System.out.println("Actual Date Range: " + actualDateRanges[k]);
            Assert.assertEquals(actualDateRanges[k], expectedDateRanges[k], "Date range is incorrect for column " + (k + 1));
        }

        System.out.println("Total/ Average: " + actualTotalAverage);
        System.out.println("Tag: " + actualTag);
        Assert.assertFalse(actualTotalAverage.isEmpty(), "Total Average column is missing or empty");
        Assert.assertFalse(actualTag.isEmpty(), "Tag column is missing or empty");
    }



}

