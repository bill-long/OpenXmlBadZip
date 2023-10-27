package com.example;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
    public static void main(String[] args) throws FileNotFoundException, IOException {

        System.out.println("Write data to an Excel Sheet");
        FileOutputStream fos = new FileOutputStream("C:\\Users\\bill\\Documents\\1.xlsx");
        SXSSFWorkbook workBook = new SXSSFWorkbook();
        SXSSFSheet spreadSheet = workBook.createSheet("email");
        SXSSFRow row;
        SXSSFCell cell;
        for (int i = 0; i < 10; i++) {
            row = spreadSheet.createRow((short) i);
            cell = row.createCell(i);

            cell.setCellValue("string value added");
        }

        workBook.write(fos);
        workBook.close();
        fos.close();
    }
}
