package com.company;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

/**
 * Performs operations with excel files.
 */
public class XFileReader {
    /**
     * Loads data from a given excel file
     *
     * @param filePath a path to the file to read from
     * @return
     */
    public XSSFWorkbook loadFromFile(final String filePath) {
        XSSFWorkbook workbook = null;
        try {
            File f = new File(filePath);
            FileInputStream file = new FileInputStream(f);
            workbook = new XSSFWorkbook(file);
            file.close();
        } catch (Exception e) {
            System.out.println("Error " + e.getMessage() + " when processing file " + filePath);
            System.out.println("Try to open the file by Excel and save it with extension .xlsx");
        }
        if (workbook != null) {
            System.out.println("Loaded " + workbook.getSheetAt(0).getPhysicalNumberOfRows() + " rows from file " + filePath);
        }
        return workbook;
    }


}
