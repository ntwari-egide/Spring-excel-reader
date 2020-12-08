package com.springreader.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
public class ExcelFormatReader {
    @GetMapping("/excelReader1")
    public String simpleReadingFromExcel()throws Exception{
        //The Apache POI library supports both .xls and .xlsx files and is a more complex library than other Java libraries for working with Excel files.
        //When working with the newer .xlsx file format, you would use the XSSFWorkbook, XSSFSheet, XSSFRow, and XSSFCell classes.
        //To work with the older .xls format, use the HSSFWorkbook, HSSFSheet, HSSFRow, and HSSFCell classes.

        File file = new File("C:\\Users\\user\\Documents\\ntwari egide documents\\spring boot\\Spring-excel-reader\\src\\main\\resources\\members.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);

        // Generating Map to store the data from the cheet
        Map<Integer , List<String>> data = new HashMap<>();
        int i = 0;

        for (Row row : sheet){
            // we start by saying first row column id because It doesn't consider the blank spaces of the sheet
            System.out.println(row.getCell(1));
            System.out.println(row.getCell(2));
            System.out.println(row.getCell(3));
            System.out.println(row.getCell(4));
            System.out.println(row.getCell(5));
        }

        return "excel file is read";
    }

    @GetMapping("/excelReader2")
    public String nextLevelOfReadingExcel(){

        return "Excel is read";
    }
}
