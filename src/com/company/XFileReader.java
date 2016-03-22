package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;

/**
 * Reads an xls file from disk
 */
public class XFileReader {

    /**
     * Number of the column in excel file to be used as an index when constructing a hash map.
     */
    private int indexColNum = 0;

    /**
     * Path to an excel file to be read
     */
    private String filePath;

    public XFileReader(String filePath, int index) {
        this.filePath = filePath;
        this.indexColNum = index;
    }


    /**
     * Loads data from a given excel file
     *
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


    /**
     * Returns true if the rows have the same content. Otherwise returns false.
     *
     * @param r1
     * @param r2
     * @return
     */
    public boolean areEqual(Row r1, Row r2) {
        int size1 = r1.getPhysicalNumberOfCells();
        if (size1 != r2.getPhysicalNumberOfCells()) {
            return false;
        }
        Cell c1, c2;
        for (int i = 0; i < size1; i++) {
            c1 = r1.getCell(i);
            c2 = r2.getCell(i);
            if (!areEqual(c1, c2)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Returns true if the cells have the same content. Otherwise returns false.
     *
     * @param c1
     * @param c2
     * @return
     */
    public boolean areEqual(Cell c1, Cell c2) {
        int t1 = c1.getCellType();
        int t2 = c2.getCellType();
        if (t1 != t2) {
            return false;
        }
        /// TODO
        return false;

    }


}
