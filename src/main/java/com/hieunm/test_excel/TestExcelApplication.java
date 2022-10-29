package com.hieunm.test_excel;

import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class TestExcelApplication {

    public static void main(String[] args) {
        // SpringApplication.run(TestExcelApplication.class, args);
        ExcelUtility excelUtility = new ExcelUtility();
        // excelUtility.createWorkbook();
        // excelUtility.readWorkbook();
        excelUtility.editWorkbook();
    }

}
