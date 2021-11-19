package com.java;
import java.io.IOException;
public class Run {
    public static void main(String[] args) throws IOException {
        ExcelHandling excelHandling= new ExcelHandling();
        {
        excelHandling.sheetName();
        excelHandling.methodToFindHeaders();
        excelHandling.methodToFindOutNumberOfRowAndColumn();
        excelHandling.methodToFindToFindTheCellContent();
        //excelHandling.getValueOfGivenRows();

        excelHandling.getValueOfGivenColumn();
        }
        ReadExcel readExcel = new ReadExcel();
        {
            readExcel.readExcelFile();
            readExcel.convertObjects2JsonString();
        }
    }
}
