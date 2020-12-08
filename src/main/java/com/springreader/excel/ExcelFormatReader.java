package com.springreader.excel;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ExcelFormatReader {
    @GetMapping("/excelReader1")
    public String simpleReadingFromExcel(){
        //The Apache POI library supports both .xls and .xlsx files and is a more complex library than other Java libraries for working with Excel files.
        //When working with the newer .xlsx file format, you would use the XSSFWorkbook, XSSFSheet, XSSFRow, and XSSFCell classes.
        //To work with the older .xls format, use the HSSFWorkbook, HSSFSheet, HSSFRow, and HSSFCell classes.


        return "excel file is read";
    }
}
