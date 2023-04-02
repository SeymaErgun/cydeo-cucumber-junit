package com.cydeo.Tests;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelRead {
@Test
public void read_from_excel_file() throws IOException {
    String path="SampleData.xlsx";
    //to read from excel we need to load it to fileInput Stream
    FileInputStream fileInputStream=new FileInputStream(path);

    //workbook>sheet>cell
    XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);

    // <2> We need to get spesific sheet from currently opened workbook
    XSSFSheet sheet=workbook.getSheet("Employees");

    //<3> Select ror and cell
    //Print out mary's cell
    //Indexes start from 0
    System.out.println(sheet.getRow(1).getCell(0));

    //print out Developer
    System.out.println(sheet.getRow(3).getCell(2));

    //Return the count of used cells only
    //Starts counting from 1
    int usedRows=sheet.getPhysicalNumberOfRows();
    System.out.println(usedRows);

    //Return the number from top cell to bottom cell
    //It doesn't care if the cell is empty or not
    //Starts counting from 0
    int lastUsedRow=sheet.getLastRowNum();
    System.out.println(lastUsedRow);

    //TODO:Create a logic to print Vinad's name
    for(int rowNum=0; rowNum<usedRows; rowNum++){
        if(sheet.getRow(rowNum).getCell(0).toString().equals("Vinod")){
            System.out.println(sheet.getRow(rowNum).getCell(0));
        }
    }

    //TODO: Create a logic to print out Linda's Job_ID
    //Check if name is Linda--->print out Job_ID of Linda
    for(int rowNum=0; rowNum<usedRows; rowNum++){
        if(sheet.getRow(rowNum).getCell(0).toString().equals("Linda")){
            System.out.println( "Linda's Job_ID is "  +sheet.getRow(rowNum).getCell(2));
        }
    }
}
}
