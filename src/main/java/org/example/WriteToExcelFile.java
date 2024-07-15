package org.example;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

public class WriteToExcelFile {
    public static void main(String[] args) throws Exception {
        File file = new File("C:\Practice\intellij_Workplace\Exam14Jul2024\src\main\resources\Employee.Excelfile.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

        int lastRowNum = xssfSheet.getLastRowNum();
        System.out.println(lastRowNum);
        Row row = xssfSheet.createRow(lastRowNum + 1);
        int noOfCell = xssfSheet.getRow(0).getLastCellNum();
        for (int i = 0; i < noOfCell; i++) {

        }
    }
}


