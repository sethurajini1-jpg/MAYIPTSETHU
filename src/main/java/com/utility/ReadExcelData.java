package com.utility;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

public class ReadExcelData {

    public static void getparticulardata(){
try{
    File file = new File("C:\\Users\\DELL\\Downloads\\DataDriven_IPT.xlsx");
    Workbook book = new XSSFWorkbook(file);
    Sheet sheet = book.getSheet("Sheet1");
    Row row = sheet.getRow(3);
    Cell cell = row.getCell(0);
    String data =cell.getStringCellValue();
    System.out.println(data);




}catch(Exception e){
    e.printStackTrace();

}


    }


    public static void main(String[]args) {
        getparticulardata();

    }

}
