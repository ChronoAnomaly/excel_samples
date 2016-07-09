/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.excel_samples;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Brett
 */
public class ExcelSamples {

    public static void main(String[] args) throws Exception {
//        FileInputStream f = new FileInputStream("schedule_template.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook("schedule_template.xlsx");

        XSSFSheet sheet = workbook.getSheet("Schedule");
//        int rowNum = 2;

        for (int rowNum = 2; rowNum < 10; rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);

            if (row == null) {
                System.out.println("Row is empty");
                row = sheet.createRow(rowNum);
            } else {
                
            }

            XSSFCell cell = row.getCell(0);

            if (cell == null) {
                System.out.println("Cell is empty");
                cell = row.createCell(0);
            } else {
                
            }
            
            cell.setCellValue("Hello");
        }

        File tmpFile = File.createTempFile("teeny_file_", ".xlsx");

        try {
            FileOutputStream out = new FileOutputStream(tmpFile);
            workbook.write(out);
            out.close();
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
