package com.java;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Scanner;

public class ReadExcel {
    List<Students>list;
    String filePath ="C:\\Users\\Udhayakumar\\IdeaProjects\\practisetask\\Excel.xlsx";
    FileInputStream inputStream;
    XSSFWorkbook workbook;XSSFSheet sheet;
    public void  readExcelFile() {
        try {
            inputStream = new FileInputStream(filePath);
            workbook= new XSSFWorkbook(inputStream);
            sheet = workbook.getSheetAt(0);
            list = new LinkedList();
            Iterator rows = sheet.iterator();
            int rowNum = 0;
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                if (rowNum == 0) {
                    rowNum++;
                    continue;

                }
                Iterator cells = row.iterator();
                Students students = new Students();
                int cellIndex = 0;
                while(cells.hasNext()){
                    XSSFCell cell= (XSSFCell) cells.next();
                    if (cellIndex ==0){
                        students.setRollNum((int) cell.getNumericCellValue());
                    }else if(cellIndex == 1){
                        students.setName(cell.getStringCellValue());
                    }else if(cellIndex == 2){
                        students.setAge((int) cell.getNumericCellValue());
                    }else if (cellIndex == 3){
                        students.setMark((int) cell.getNumericCellValue());
                    }
                    cellIndex ++;
                }
                list.add(students);
                            }
            workbook.close();
            System.out.println(list);
        }catch (IOException e){
            System.out.println("Alert !!!! File not found in the given path");
        }
    }

    public void convertObjects2JsonString() throws IOException {
        File file;
        Scanner scan = new Scanner(System.in);
        System.out.println("Name Your File with .json");
        String filename = scan.nextLine();
        String directory = System.getProperty("user.dir");
        String path;
        path = directory + File.separator + filename;
        System.out.println("Filepath is " + path);
        file = new File(path);
        ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.writerWithDefaultPrettyPrinter().writeValue(file,list);
        System.out.println(list);
    }
}
