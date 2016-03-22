package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;

/**
 * Performs operations with excel files.
 */
public class XFileReader {
    /**
     * Loads data from a given excel file
     * @param filePath a path to the file to read from
     * @return
     */
    public XSSFWorkbook loadFromFile(final String filePath) {
        XSSFWorkbook workbook = null;
        try {
            FileInputStream file = new FileInputStream(new File(filePath));
            workbook = new XSSFWorkbook(file);
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("Loaded " + workbook.getSheetAt(0).getPhysicalNumberOfRows() + " rows from file " + filePath);
        return workbook;
    }





}
