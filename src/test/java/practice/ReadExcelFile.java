package practice;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;


public class ReadExcelFile {

    public ReadExcelFile(){

    }

    public void readExcel(String filePath, String sheetName) throws IOException {

        File file = new File(filePath);
        FileInputStream inputStream = new FileInputStream(file);

        XSSFWorkbook newWorkBook = new XSSFWorkbook(inputStream);
        XSSFSheet newSheet = newWorkBook.getSheet(sheetName);

        int rowCount = newSheet.getLastRowNum() - newSheet.getFirstRowNum();

        // Vamos iterando por cada fila
        for (int i = 0; i<=rowCount; i++ ) {
            XSSFRow row = newSheet.getRow(i);

            // Vamos iterando por las celdas de dicha fila
            for (int j = 0; j<row.getFirstCellNum(); j++) {
                System.out.println(row.getCell(j).getStringCellValue() + "||");
            }
        }
    }

    public String getCellValue(String filePath, String sheetName, int rowNumber, int cellNumber) throws IOException {

        File file = new File(filePath);
        FileInputStream inputStream = new FileInputStream(file);

        XSSFWorkbook newWorkBook = new XSSFWorkbook(inputStream);
        XSSFSheet newSheet = newWorkBook.getSheet(sheetName);

        XSSFRow row = newSheet.getRow(rowNumber);
        XSSFCell cell = row.getCell(cellNumber);

        return cell.getStringCellValue();
    }

}
