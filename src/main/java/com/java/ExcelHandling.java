package com.java;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

public class ExcelHandling extends ReadExcel {
    Scanner input=new Scanner(System.in);
    DataFormatter dataFormatter = new DataFormatter();
    public ExcelHandling(){

        try{ inputStream = new FileInputStream(filePath);
            workbook = new XSSFWorkbook(inputStream);}
        catch (Exception e){
            e.printStackTrace();
        }
    }
    public void sheetName(){
        System.out.println(filePath);
        Iterator<Sheet>sheetIterator=workbook.sheetIterator();
        while (sheetIterator.hasNext()){
            Sheet sheetName = sheetIterator.next();
            System.out.println("Sheet name =====>" +sheetName.getSheetName());
            System.out.println("---------------------------------------------------");
    }
    }

    public void methodToFindHeaders(){
        sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        System.out.println("The Sheet with following Headers");
        for (int i = 0 ; i<=3 ; i++) {
            XSSFCell cell = row.getCell(i);
            String header=cell.getStringCellValue();
            System.out.print("\""+header+ "\"\t");
        }
        System.out.println("\n---------------------------------------------------");
    }
    public void methodToFindOutNumberOfRowAndColumn(){
        sheet = workbook.getSheetAt(0);
        Iterator rowIterator = sheet.rowIterator();
        int numberOfRow = sheet.getLastRowNum();
        int physicalNumberOfRows =sheet.getPhysicalNumberOfRows();
        System.out.println("Number of rows are " +numberOfRow+ " without Header");
        System.out.println("Number of rows are " +physicalNumberOfRows+ " with header");
        System.out.println("---------------------------------------------------");
        int numberOfColumn =0;
        if (rowIterator.hasNext())
        {
            Row headerRow = (Row) rowIterator.next();
            numberOfColumn = headerRow.getPhysicalNumberOfCells();
        }
        System.out.println("number of column "+ numberOfColumn);
        System.out.println("---------------------------------------------------");

    }

    public void getValueOfGivenRows() throws IOException {
        sheet = workbook.getSheetAt(0);
        System.out.println("Enter the row value");
        int rowValue = input.nextInt();
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row rowData = sheet.getRow(rowValue);
        Iterator<Cell> cellIterator = rowData.cellIterator();
        System.out.println("the data of the row "+rowValue+ " are");
        while (rowIterator.hasNext())
        {
            while (cellIterator.hasNext())
            {
                Cell cell = cellIterator.next();
                System.out.println(cell);
            }
        }
        }
        public void getValueOfGivenColumn(){
        sheet = workbook.getSheetAt(0);
        System.out.println("Enter the column Index");
        int columnIndex= input.nextInt();
            Iterator<Row>rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                   if(cell.getColumnIndex()==columnIndex)
                    {
                        System.out.println(cell);
                    }

        }
    }
    }
    public void methodToFindToFindTheCellContent(){

        System.out.println("Enter the row Value");
        int i = input.nextInt();
        System.out.println("Enter the cell value");//it will take like(0,0)
        int j = input.nextInt();
        sheet = workbook.getSheetAt(0);
        XSSFRow row =sheet.getRow(i);
        Cell cell=row.getCell(j);
        String value = dataFormatter.formatCellValue(cell);
        System.out.println("The Cell contains " +value);
        System.out.println("---------------------------------------------------");

    }
}