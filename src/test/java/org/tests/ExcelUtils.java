/*
package org.tests;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelUtils {
  private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
  static ExcelUtils excel;
  private static HSSFSheet ExcelWSheet;
  private static HSSFWorkbook ExcelWBook;
  private static HSSFCell Cell;
  private static HSSFRow Row;

  */
/**
   * Return table array for an given Excel file and sheet
   *
   * @param FilePath
   * @param SheetName
   * @return
   * @throws Exception
   *//*

  public static Object[][] getTableArray(String FilePath, String SheetName) throws Exception {
    String[][] tabArray = null;

    try {
      String extension =
          FilenameUtils.getExtension(FilePath); // getting file extension from string given
      // System.out.println(extension);
      FileInputStream ExcelFile = new FileInputStream(FilePath);

      //			if(extension.contentEquals("xlsx"))
      //			{
      // Access the required test data sheet
      ExcelWBook = new HSSFWorkbook(ExcelFile);
      ExcelWSheet = ExcelWBook.getSheet(SheetName);
      int startRow = 1;
      int startCol = 0;
      int ci, cj;
      int totalRows = ExcelWSheet.getLastRowNum();
      //			   System.out.println("TotalRows: "+totalRows);
      // you can write a function as well to get Column count
      int totalCols = ExcelWSheet.getRow(0).getPhysicalNumberOfCells();
      // int totalCols = ExcelWSheet.getRow(i).getLastCellNum() ;
      System.out.println("Total Columns :" + totalCols);
      tabArray = new String[totalRows][totalCols];
      ci = 0;
      for (int i = startRow; i <= totalRows; i++, ci++) {
        cj = 0;
        for (int j = startCol; j < totalCols; j++, cj++) {
          tabArray[ci][cj] = getCellData(i, j);
          // System.out.println(tabArray[ci][cj]);
        }
      }
    } catch (FileNotFoundException e) {
      logger.error("Could not read the Excel sheet", e);
      e.printStackTrace();
    } catch (IOException e) {

      logger.error("Could not read the Excel sheet", e);
      e.printStackTrace();
    }

    return (tabArray);
  }

  */
/**
   * Return cell data an given row and column from an specified Excel
   *
   * @param RowNum
   * @param ColNum
   * @return
   * @throws Exception
   *//*

  public static String getCellData(int RowNum, int ColNum) throws Exception {

    try {
      int dataType = 0;
      Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
      String CellData = null;
      if (Cell == null) {
        dataType = 3;
      } else {
        dataType = Cell.getCellType();
      }

      if (dataType == 3) {
        return "";
      } else if (dataType == 0) {
        if (HSSFDateUtil.isCellDateFormatted(Cell)) {
          DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
          CellData = df.format(Cell.getDateCellValue());

        } else {
          double doubleCellData = Cell.getNumericCellValue();
          CellData = String.valueOf((int) doubleCellData);
        }
        return String.valueOf(CellData);
      } else {
        CellData = Cell.getStringCellValue();
        return CellData;
      }
    } catch (Exception e) {
      System.out.println(e.getMessage());
      throw (new Exception("Not able to read excel", e));
    }
  }
  */
/**
   * Return Cell value for an given Excel file details
   *
   * @param xlpath
   * @param sheetName
   * @param rowNum
   * @param cellNum
   * @return
   * @throws FileNotFoundException
   * @throws IOException
   * @throws IllegalStateException
   * @throws InvalidFormatException
   *//*

  public static String getExcelCellvalue(String xlpath, String sheetName, int rowNum, int cellNum)
      throws FileNotFoundException, IOException, IllegalStateException, InvalidFormatException {
    // TODO Auto-generated method stub
    String cellVal;

    try {
      FileInputStream fis = new FileInputStream(xlpath);
      Workbook wb = WorkbookFactory.create(fis);
      cellVal = wb.getSheet(sheetName).getRow(rowNum).getCell(cellNum).getStringCellValue();
      //			System.out.println(cellVal);
    } catch (Exception e) {
      cellVal = "";
    }
    return cellVal;
  }

  */
/**
   * Method to read all the values to an string array by passing a particular row name
   *
   * @param xlPath
   * @param xlSheet
   * @param rowName
   * @return
   * @throws Exception
   *//*

  public static String[] getExcelRowValues(String xlPath, String xlSheet, String rowName)
      throws Exception {
    String[] detailsArray = null;
    try {
      FileInputStream ExcelFile = new FileInputStream(xlPath);
      // Access the required test data sheet
      ExcelWBook = new HSSFWorkbook(ExcelFile);
      ExcelWSheet = ExcelWBook.getSheet(xlSheet);
      int totalRows = ExcelWSheet.getLastRowNum();
      int ci = 0, totalCols = 0;
      int rowNum = 0;
      for (int i = 1; i <= totalRows; i++) {
        //				String strExcelRow=ExcelWSheet.getRow(i).getCell(0).getStringCellValue();
        String strExcelRow = getCellData(i, 0);
        if (strExcelRow.equalsIgnoreCase(rowName)) {
          rowNum = i;
        }
      }

      if (rowNum != 0) {
        totalCols = ExcelWSheet.getRow(rowNum).getLastCellNum();
        detailsArray = new String[totalCols];
        for (int j = 0; j < totalCols; j++, ci++) {
          detailsArray[ci] = getCellData(rowNum, j);
          // detailsArray[ci] =	ExcelWSheet.getRow(rowNum).getCell(j).getStringCellValue();
        }
      } else {
        logger.info("Expected row not found in the excel");
      }
    } catch (IOException | NullPointerException e) {
      logger.error("Could not read the Excel sheet", e);
      throw (new Exception("Could not read the Excel sheet, please check it ", e));
    }

    return (detailsArray);
  }

  */
/**
   * Method to read all the values to an string array by passing a particular column name
   *
   * @param xlPath
   * @param xlSheet
   * @param columnName
   * @return
   * @throws Exception
   *//*

  public static String[] getExcelColumnValues(String xlPath, String xlSheet, String columnName)
      throws Exception {
    String[] arrColumnValues = null;
    try {
      FileInputStream ExcelFile = new FileInputStream(xlPath);
      // Access the required test data sheet
      ExcelWBook = new HSSFWorkbook(ExcelFile);
      ExcelWSheet = ExcelWBook.getSheet(xlSheet);
      int totalRows = ExcelWSheet.getLastRowNum();

      int ci = 0, colNum = 0;
      int rowNum = 0;
      for (int i = 0; i < ExcelWSheet.getRow(0).getPhysicalNumberOfCells(); i++) {
        String strColName = getCellData(0, i);
        if (strColName.contains(columnName)) {
          colNum = i;
          break;
        } else {
          colNum = -1;
        }
      }

      if (colNum != -1) {
        // totalCols=ExcelWSheet.getRow(rowNum).getLastCellNum();
        arrColumnValues = new String[totalRows];
        for (int j = 1; j <= totalRows; j++, ci++) {
          arrColumnValues[ci] = getCellData(j, colNum);
        }
      } else {
        logger.error("Expected column not found in the excel");
      }
    } catch (IOException | NullPointerException e) {
      logger.error("Could not read the Excel sheet", e);
      throw new FrameworkException("Could not read the Excel sheet, please check it", e);
    }

    return (arrColumnValues);
  }

  // method to get values from excel from a specific sheet
  */
/**
   * Method to get values from excel from a specific sheet
   *
   * @param FilePath
   * @param SheetName
   * @return
   * @throws Exception
   *//*

  public static String[][] getValuesFrmExcel(String FilePath, String SheetName) throws Exception {
    String[][] arrExcelValues = null;
    excel = new ExcelUtils();
    try {

      FileInputStream ExcelFile = new FileInputStream(FilePath);
      // Access the required test data sheet
      ExcelWBook = new HSSFWorkbook(ExcelFile);
      ExcelWSheet = ExcelWBook.getSheet(SheetName);
      int startRow = 1;
      int startCol = 0;
      int ci, cj;
      int totalRows = ExcelWSheet.getLastRowNum();
      int totalCols = ExcelWSheet.getRow(0).getPhysicalNumberOfCells();
      arrExcelValues = new String[totalRows][totalCols];
      // int totalCols=2;
      ci = 0;
      for (int i = startRow; i <= totalRows; i++, ci++) {
        cj = 0;
        for (int j = startCol; j < totalCols; j++, cj++) {
          arrExcelValues[ci][cj] = excel.getCellData(i, j);
        }
      }
    } catch (FileNotFoundException e) {
      logger.error("Exception while reading the excel", e);
      throw new FrameworkException("Exception while reading the excel", e);
    }

    return (arrExcelValues);
  }

  // TODO need to check the method
  */
/*
   * Method to launch website with an given URL
   *//*

  */
/*   public void NavigateTo(String URL) {
   *//*

  */
/*    try{
          driver.navigate().to(URL);
          logger.log(LogStatus.INFO, "Launching URL: "+URL);
      }
      catch(Exception e){
          getScreenshot(driver,"Error occured while launching website.",false);

      }
  }*//*

  */
/*
  }*//*


  */
/**
   * Method to get row values into string array
   *
   * @param xlpath
   * @param sheetName
   * @param rowNum
   * @return String[] consists of the row values
   * @throws FrameworkException
   *//*

  public static String[] getExcelRowValues(String xlpath, String sheetName, int rowNum)
      throws Exception {
    String[] rowValues = null;
    try {
      FileInputStream ExcelFile = new FileInputStream(xlpath);
      ExcelWBook = new HSSFWorkbook(ExcelFile);
      ExcelWSheet = ExcelWBook.getSheet(sheetName);
      // The below line always fetch the total column in the Header
      int columnCnt = ExcelWSheet.getRow(0).getLastCellNum();
      rowValues = new String[columnCnt];
      for (int i = 0; i < columnCnt; i++) {
        rowValues[i] = getCellData(rowNum, i);
      }
    } catch (IOException e) {
      logger.error(
          "Error occured while reading excel row values, error message is: " + e.getMessage());
      throw new FrameworkException("Exception while reading the excel row value: ", e);
    }
    return rowValues;
  }

  */
/**
   * Method to get total rows present in the excel
   *
   * @param xlpath
   * @param sheetName
   * @param excluedColHeader: if true first row is not considered, if false first row is considered
   * @return int (Total Row count)
   *//*

  public static int getRowCount(String xlpath, String sheetName, boolean excluedColHeader) {
    int rowCount = 0;
    try {
      FileInputStream ExcelFile = new FileInputStream(xlpath);
      ExcelWBook = new HSSFWorkbook(ExcelFile);
      int totalRows = ExcelWBook.getSheet(sheetName).getPhysicalNumberOfRows();
      if (excluedColHeader) {
        rowCount = totalRows - 1;
      } else {
        rowCount = totalRows;
      }
    } catch (FileNotFoundException file) {
      logger.error("Specified Excel File not found, File path: " + xlpath);
      rowCount = 0;
    } catch (Exception e) {
      logger.error("Error occured while getting row count,error message is :" + e.getMessage());
      rowCount = 0;
    }
    return rowCount;
  }

  */
/**
   * Method to read the specific row value as HashMap.
   *
   * @param filePath
   * @param sheetName
   * @param rowName
   * @return HashMap
   * @throws FrameworkException
   *//*

  public static HashMap<String, String> getExcelRowValue(
      String filePath, String sheetName, String rowName) throws FrameworkException {
    int numberOfCells = 0;
    // Used the LinkedHashMap to maintain the order
    HashMap<String, String> output = new LinkedHashMap<String, String>();

    FileInputStream fis = null;
    try {
      fis = new FileInputStream(filePath);
      ExcelWBook = new HSSFWorkbook(fis);
      ExcelWSheet = ExcelWBook.getSheet(sheetName);
      numberOfCells = ExcelWSheet.rowIterator().next().getPhysicalNumberOfCells();

      String[] headerValues = getExcelRowValues(filePath, sheetName, 0);
      String[] rowValues = getExcelRowValues(filePath, sheetName, rowName);

      for (int i = 0; i < numberOfCells; i++) {
        output.put(headerValues[i], rowValues[i]);
      }
    } catch (Exception f) {
      logger.error(
          "Error occured while reading the excel data, Method Name: getExcelRowValue.Error message is: "
              + f.getMessage());
      throw new FrameworkException(
          "Exception while reading the excel row value, Method Name: getExcelRowValue: ", f);
    } finally {
      if (fis != null) {
        try {
          fis.close();
        } catch (IOException e) {
          logger.error(
              "Error occured while closing the excel file, Method Name: getExcelRowValue.Error message is: "
                  + e.getMessage());
          throw new FrameworkException(
              "Exception while reading the excel row value, Method Name: getExcelRowValue: ", e);
        }
      }
    }
    return output;
  }

  */
/**
   * Method to read data sheet values as Hash map This method will find the matching key in the
   * first column (Unique Key) values and get the row values from the matched rows
   *
   * @param filePath
   * @param sheetName
   * @param Key
   * @return Hash map
   * @throws Exception
   *//*

  public static Object[][] getSheetData(String filePath, String sheetName, String UniqueKey)
      throws Exception {

    FileInputStream fis = null;
    Object[][] obj = null;
    Map<Object, Object> datamap = null;

    try {
      fis = new FileInputStream(filePath);

      ExcelWBook = new HSSFWorkbook(fis);
      ExcelWSheet = ExcelWBook.getSheet(sheetName);
      ExcelWBook.close();
      // int lastRowNum = ExcelWSheet.getLastRowNum() ;
      int lastCellNum = ExcelWSheet.getRow(0).getLastCellNum();
      int matchingRowCount = 0;
      int count = 0;

      String[] colValues = getExcelColumnValues(filePath, sheetName, "Unique Key");

      // Getting the matched row counts
      for (int i = 0; i < colValues.length; i++) {
        if (colValues[i].contains(UniqueKey)) {
          matchingRowCount++;
        }
      }

      obj = new Object[matchingRowCount][1];

      // Building Hash Map
      for (int i = 0; i < colValues.length; i++) {
        if (colValues[i].contains(UniqueKey)) {
          int rowIndex = getIndexOf(colValues, colValues[i]);
          String[] rowValues = getExcelRowValues(filePath, sheetName, rowIndex + 1);
          datamap = new LinkedHashMap<>();
          for (int j = 0; j < lastCellNum; j++) {
            datamap.put(getCellData(0, j).toString(), rowValues[j]);
          }
          obj[count][0] = datamap;
          count++;
        } else {
          continue;
        }
      }
      */
/*
      for (int i= 0; i< lastRowNum; i++) {
      	if(getCellData(i+1, 0).contains(UniqueKey)) {
      		datamap = new LinkedHashMap<>();
      		//PickNextRow:
      		for (int j= 0; j< lastCellNum; j++) {
      			//datamap.put(sheet.getRow(0).getCell(j).toString(), sheet.getRow(i+1).getCell(j).toString());

      			//Validating the Key value to pick the row values
      			//if(getCellData(i+1, 0).contains(UniqueKey)) {
      			datamap.put(getCellData(0, j).toString(), getCellData((i+1), j).toString());
      			//}else {
      			//	break PickNextRow;
      			//	}
      		}
      		obj[i][0] = datamap;
      	}else {
      		continue;
      	}
      }
       *//*
 } catch (Exception f) {
      logger.error(
          "Error occured while reading the excel data- Method Name: getSheetData. Error message is: "
              + f.getMessage());
      throw new FrameworkException(
          "Exception while reading the excel row value- Method Name: getSheetData.: ", f);
    } finally {
      if (fis != null) {
        try {
          fis.close();
        } catch (IOException e) {
          logger.error(
              "Error occured while closing the excel file- Method Name: getSheetData.Error message is: "
                  + e.getMessage());
          throw new FrameworkException(
              "Exception while reading the excel row value- Method Name: getSheetData.: ", e);
        }
      }
    }

    return obj;
  }

  */
/*
   * Method to Get the Index value from the String Array
   * @Parameters= String array, Value for which the Index is required
   * Return's Index of the matching Value in the Array
   * Returns -1 if the value doesn't found in the array
   * NOTE: Index is of Array Type.
   *//*

  */
/**
   * Method to get index value from an String array values
   *
   * @param arrayValues
   * @param value
   * @return
   *//*

  public static int getIndexOf(String[] arrayValues, String value) {
    int retValue = -1;
    try {
      for (int i = 0; i < arrayValues.length; i++) {
        if (arrayValues[i].equals(value)) {
          retValue = i; // Matched
          break;
        } else {
          retValue = -1; // Not Matched
          continue;
        }
      }
    } catch (Exception e) {
      logger.info("Error occured while getting the Index of - " + value);
    }
    return retValue;
  }
}
*/
