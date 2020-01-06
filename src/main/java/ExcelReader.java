import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;

    public static void setExcelFile(String Path, String SheetName) throws Exception {

        try {
            // Open the Excel file
            FileInputStream ExcelFile = new FileInputStream(Path);
            // Access the required test data sheet
            ExcelWBook = new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);
        } catch (Exception e) {
            throw (e);
        }
    }

    public static Object[][] getTableArray() throws Exception {

        String[][] tabArray = null;
        try {

            int ci = 0, cj = 0;
            int startRow = ExcelWSheet.getFirstRowNum(), startCol = ExcelWSheet.getRow(ExcelWSheet.getFirstRowNum()).getFirstCellNum();
            int totalRows = ExcelWSheet.getLastRowNum() - ExcelWSheet.getFirstRowNum();
            int totalCols = ExcelWSheet.getRow(ExcelWSheet.getFirstRowNum()).getLastCellNum() - ExcelWSheet.getRow(ExcelWSheet.getFirstRowNum()).getFirstCellNum();
            tabArray = new String[totalRows+1][totalCols+1];

            for (int i = startRow; i <= totalRows; i++, ci++) {
                for (int j = startCol; j <= totalCols; j++, cj++) {
                    tabArray[ci][cj] = getCellData(i, j);
                }
                cj=0;
            }
        } catch (FileNotFoundException e) {
            System.out.println("Could not read the Excel sheet");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("Could not read the Excel sheet");
            e.printStackTrace();
        }
        return (tabArray);
    }

    public static String fetchData(String rowName, String columnName) {

        try{
            Object dArray[][] = getTableArray();
            int rowIndex = 0, colIndex = 0;

            for (int i = 0; i < dArray[0].length ; i++) {
                if(dArray[0][i].equals(columnName)) {
                    colIndex = i;
                    break;
                }
            }
            for (int j = 0; j < dArray.length ; j++) {
                if(dArray[j][0].equals(rowName)) {
                    rowIndex = j;
                    break;
                }
            }
            return (String) dArray[rowIndex][colIndex];
        } catch (Exception e) {
            System.out.println("Exception Occurred. More information below");
            e.printStackTrace();
            return null;
        }
    }

    public static String getCellData(int RowNum, int ColNum) throws Exception {

        try {
            Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
            String CellData = Cell.getStringCellValue();
            return CellData;
        } catch (Exception e) {
            return null;
        }
    }

    public static String getTestCaseName(String sTestCase) throws Exception {

        String value = sTestCase;
        try {
            int posi = value.indexOf("@");
            value = value.substring(0, posi);
            posi = value.lastIndexOf(".");
            value = value.substring(posi + 1);
            return value;
        } catch (Exception e) {
            throw (e);
        }
    }

    public static int getRowContains(String rowName) throws Exception {

        int i;
        try {
            int rowCount = ExcelWSheet.getLastRowNum();
            for (i = 0; i < rowCount; i++) {
                if (getCellData(i, 0).equalsIgnoreCase(rowName)) {
                    break;
                }
            }
            return i;
        } catch (Exception e) {
            throw (e);
        }
    }


    public void readExcel(String filePath, String fileName, String sheetName) throws IOException {
        File file = new File(filePath + "\\" + fileName);
        FileInputStream inputStream = new FileInputStream(file);
        Workbook oWorkBook = null;

        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        if (fileExtensionName.equals(".xlsx")) {
            oWorkBook = new XSSFWorkbook(inputStream);
        } else if (fileExtensionName.equals(".xls")) {
            oWorkBook = new HSSFWorkbook(inputStream);
        }

        Sheet oSheet = oWorkBook.getSheet(sheetName);
        int rowCount = oSheet.getLastRowNum() - oSheet.getFirstRowNum();

        for (int i = 0; i < rowCount + 1; i++) {
            Row row = oSheet.getRow(i);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                System.out.print(row.getCell(j).getStringCellValue() + "|| ");
            }
            System.out.println();
        }
    }

    public static void writeExcel(String path, String sheetName, String rowName, String newColName, String newData) {

        try {
            FileInputStream excelFile = new FileInputStream(new File(path));
            XSSFWorkbook oBook = new XSSFWorkbook(excelFile);
            XSSFSheet oSheet = oBook.getSheet(sheetName);

            String[] newCol = newColName.split("-##-");
            String[] newData

            for (int c = 0; c < newCol.length; c++) {

            }

            int startRow = oSheet.getFirstRowNum(), startCol = oSheet.getRow(oSheet.getFirstRowNum()).getFirstCellNum();
            int totalRows = oSheet.getLastRowNum() - oSheet.getFirstRowNum();
            int totalCols = oSheet.getRow(oSheet.getFirstRowNum()).getLastCellNum() - oSheet.getRow(oSheet.getFirstRowNum()).getFirstCellNum();
            int rowIndex = 0, newColIndex;
            Cell temp = oSheet.getRow(startRow).getCell(oSheet.getRow(startRow).getLastCellNum() - 1);
            if(temp.getStringCellValue().equals(newColName)) {
                newColIndex = oSheet.getRow(startRow).getLastCellNum() - 1;
            } else {
                newColIndex = oSheet.getRow(startRow).getLastCellNum();
            }
            Row headerRow = oSheet.getRow(oSheet.getFirstRowNum());
            Cell colCell = headerRow.createCell(newColIndex);
            colCell.setCellValue(newColName);

            for (int r = startRow ; r <= totalRows; r++) {
                int c = oSheet.getRow(startRow).getFirstCellNum();
                if(oSheet.getRow(r).getCell(c).getStringCellValue().equals(rowName)) {
                    rowIndex = r;
                    break;
                }
            }

            Row dataRow = oSheet.getRow(rowIndex);
            Cell dataCell = dataRow.createCell(newColIndex);
            dataCell.setCellValue(newData);
            excelFile.close();

            FileOutputStream excelOutput = new FileOutputStream(new File(path));
            oBook.write(excelOutput);
            oBook.close();
            excelOutput.flush();
            excelOutput.close();

        } catch (IOException i) {
            System.out.println("File not found.");
        }
    }
}
