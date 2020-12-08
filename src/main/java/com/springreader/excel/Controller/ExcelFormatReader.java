package com.springreader.excel.Controller;

import com.springreader.excel.Model.Member;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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
            if(i == 0){
                i++;
            }
            else {
                System.out.println(row.getCell(1));
                System.out.println(row.getCell(2));
                System.out.println(row.getCell(3));
                System.out.println(row.getCell(4));
                System.out.println(row.getCell(5));

                i++;
            }
        }

        return "excel file is read";
    }

    @GetMapping("/excelReader2")
    public List<Member> nextLevelOfReadingExcel()throws Exception{
        File file = new File("C:\\Users\\user\\Documents\\ntwari egide documents\\spring boot\\Spring-excel-reader\\src\\main\\resources\\members.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

//        Map<Integer,List<String>> dataFound = new HashMap<>();
        List<Member> members = new ArrayList<>();

        int i=0;

        for (Row row: sheet){
            if(i == 0){
                // to skip the header of the table in excel
                i++;
            }
            else{
                DataFormatter dataFormatter = new DataFormatter();
                String memberId = dataFormatter.formatCellValue(row.getCell(1));
                String memberFirstName = dataFormatter.formatCellValue(row.getCell(2));
                String memberSecondName =dataFormatter.formatCellValue(row.getCell(3));
                String memberMarks = dataFormatter.formatCellValue(row.getCell(4));
                String memberPosition = dataFormatter.formatCellValue(row.getCell(5));

                Member member = new Member(memberId,memberFirstName,memberSecondName,memberMarks,memberPosition);

                members.add(member);
                i++;
            }
        }

        return members;
    }

    @PostMapping("/createExcel")
    public ResponseEntity<?> createExcelFile(@RequestBody List<Member> members)throws Exception{
        System.out.println(members);
        String[] columns = {"member no", "first name", "second name", "marks","position"};

        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Employee");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        int rowNum = 1;
        for (Member member: members){
            Row row = sheet.createRow(rowNum++);
            row.createCell(1)
                    .setCellValue(member.getMemberNo());
            row.createCell(2)
                    .setCellValue(member.getFirstName());
            row.createCell(3)
                    .setCellValue(member.getSecondName());
            row.createCell(4)
                    .setCellValue(member.getMarks());
            row.createCell(5)
                    .setCellValue(member.getPosition());
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("created-excel-file.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

        return ResponseEntity.ok().build();
    }
}
