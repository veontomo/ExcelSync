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
        System.out.println(workbook.getSheetAt(0).getPhysicalNumberOfRows());
        return workbook;
    }

    /**
     * Create an index of given workbook: a map from string content of cells of given column to the the number of row
     * in which that string is found.
     *
     * @param workbook
     * @param column
     * @return
     */
    public HashMap<String, Integer> index(final XSSFWorkbook workbook, final int column) throws Exception {
        HashMap<String, Integer> map = new HashMap<>();
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rowsNum = sheet.getPhysicalNumberOfRows();
        Row row;
        Cell cell;
        String key;
        for (int i = 0; i < rowsNum; i++) {
            row = sheet.getRow(i);
            cell = row.getCell(column);
            if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
                throw new Exception("Cell " + column + " at row " + i + " is not of string type!");
            }
            key = cell.getStringCellValue();
            if (map.containsKey(key)) {
                throw new Exception("Duplicate key: " + key);
            }
            map.put(key, i);
        }
        return map;
    }



}
