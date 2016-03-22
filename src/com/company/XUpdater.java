package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * Updates a workbook with data from other workbooks
 */
public class XUpdater {

    private HashMap<String, Integer> targetIndex;
    private List<HashMap<String, Integer>> sourcesIndex;

    private final XSSFWorkbook target;
    private final XSSFWorkbook[] sources;


    /**
     * list of keys that are present in both targetIndex and sourcesIndex
     */

    private HashMap<String, Integer> duplicates;
    /**
     * list of keys that are present in targetIndex and not present in any of the sourcesIndex
     */
    private List<String> missing;

    /**
     * list of keys that are present in one of the sourcesIndex and not present in the targetIndex
     */
    private HashMap<String, Integer> extra;

    public XUpdater(final XSSFWorkbook workbook, final XSSFWorkbook[] workbooks) {
        this.target = workbook;
        this.sources = workbooks;
    }


    /**
     * Finds the sourcesIndex that contain a key from the targetIndex.
     * <p>
     * If a key is found in multiple sourcesIndex, an exception is thrown.
     * <p>
     * Returns a hash map from a string to an integer that is the ordinal number of the source in the source list in which
     * the key is found.
     *
     * @return
     */
    public void analyze() throws Exception {
        duplicates = new HashMap<>();
        missing = new ArrayList<>();
        extra = new HashMap<>();
        sourcesIndex = new ArrayList<>();

        targetIndex = index(target, 1);
        int sourcesLen = sources.length;

        for (int i = 0; i < sourcesLen; i++){
            sourcesIndex.add(index(sources[i], 0));
        }


        boolean isFoundInSources;
        // first pass: iterate over the targetIndex and control the presence in the sourcesIndex
        for (String key : targetIndex.keySet()) {
            isFoundInSources = false;
            for (int i = 0; i < sourcesLen; i++) {
                if (sourcesIndex.get(i).containsKey(key)) {
                    if (duplicates.containsKey(key)) {
                        throw new Exception("key " + key + " has already been found in source n. " + i + ". Resolve to proceed.");
                    }
                    isFoundInSources = true;
                    duplicates.put(key, i);
                }
            }
            if (!isFoundInSources) {
                missing.add(key);
            }
        }
        // second pass: iterate over the sourcesIndex and control if they contain keys that are not in the targetIndex
        for (int i = 0; i < sourcesLen; i++) {
            for (String key : sourcesIndex.get(i).keySet()) {
                if (targetIndex.containsKey(key)) {
                    // cross check: the variable "duplicates" must contain this key as well.
                    if (duplicates.containsKey(key) && duplicates.get(key) == i) {
//                        System.out.println("cross-check is OK");
                    } else {
                        System.out.println("cross-check is not OK for key " + key + " that is supposed to be in set " + i);
                    }
                } else {
                    extra.put(key, i);
                }
            }
        }

    }

    public HashMap<String, Integer> getDuplicates() {
        return duplicates;
    }

    public List<String> getMissing() {
        return missing;
    }

    public HashMap<String, Integer> getExtra() {
        return extra;
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
