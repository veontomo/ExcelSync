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
    public HashMap<String, Row> loadFromFile() {
        HashMap<String, Row> data = new HashMap<>();
        try {
            FileInputStream file = new FileInputStream(new File(filePath));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            String key;
            Row row;
            Cell indexCell;
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                indexCell = row.getCell(indexColNum);
                switch (indexCell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        key = indexCell.getStringCellValue().trim();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        key = String.valueOf(indexCell.getNumericCellValue()).trim();
                        break;
                    default:
                        throw new Exception("Non supported cell type.");
                }
//                key = row.getCell(indexColNum).getStringCellValue();

                if (data.containsKey(key)) {
                    throw new Exception("duplicate key: " + key + " in file " + this.filePath);
                } else {
                    data.put(key, row);
                }

            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return data;
    }

    /**
     * Create an index of given workbook: a map from string content of cells of given column to the the number of row
     * in which that string is found.
     * @param workbook
     * @param column
     * @return
     */
    public  HashMap<String, Integer> index(XSSFWorkbook workbook, int column){
       return null;
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
