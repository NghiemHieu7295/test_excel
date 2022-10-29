package com.hieunm.test_excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelUtility {
    public void createWorkbook() {
        Workbook wb = new XSSFWorkbook();
        Sheet firstSheet = wb.createSheet("First");

        try {
            OutputStream os = new FileOutputStream("static_content\\test.xlsx");
            wb.write(os);
            os.close();
            wb.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    public void readWorkbook() {
        try {
            Workbook wb = WorkbookFactory.create(new File("static_content\\test.xlsx"));
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(0);
            if (row == null) {
                row = sheet.createRow(0);
            }
            Cell cell = row.getCell(0);
            if (cell == null) {
                cell = row.createCell(0);
            }

            String s = "Nothing";
            if (!"".equals(cell.getStringCellValue())) {
                s = cell.getStringCellValue();
            }
            System.out.println(s);

            wb.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void editWorkbook() {
        try {
            // Open excel file
            InputStream inputStream = new FileInputStream("static_content\\test.xlsx");
            Workbook workBook = WorkbookFactory.create(inputStream);

            // Edit sheet's content
            Sheet sheet = workBook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            Row lastDataRow = sheet.getRow(lastRowNum - 2);
            CellStyle style = lastDataRow.getCell(0).getCellStyle();

            // Delete empty rows
            // sheet.shiftRows(lastRowNum, lastRowNum, -9);
            // Insert empty rows
            sheet.shiftRows(lastRowNum - 1, lastRowNum, 3);

            lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i < lastRowNum - 1; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    row = sheet.createRow(i);

                    for (int j = 0; j < 3; j++) {
                        Cell emptyCell = row.createCell(j);
                        emptyCell.setCellStyle(style);
                    }
                }
            }

            // Save file
            OutputStream outputStream = new FileOutputStream("static_content\\test.xlsx");
            workBook.write(outputStream);

            // Close stream and file
            outputStream.close();
            workBook.close();
            inputStream.close();

            System.out.println("Edit success!");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
