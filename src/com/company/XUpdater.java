package com.company;

import com.sun.istack.internal.NotNull;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
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

    private final XSSFWorkbook[] sources;
    private final int sourcesLen;
    private final XSSFWorkbook target;
    private final HashMap<Integer, Integer> map;

    private final int targetIndexCol;
    private final int sourceIndexCol;

    private HashMap<String, Integer> targetIndex;
    private List<HashMap<String, Integer>> sourcesIndex;


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

    /**
     * Constructor.
     *
     * @param workbook       a target workbook
     * @param workbooks      array of source workbooks
     * @param targetIndexCol a number of the column of the target workbook w.r.t. which an index is to be constructed
     * @param sourceIndexCol a number of the column of the source workbooks w.r.t. which an index is to be constructed
     * @param map            defines the mapping from the target workbook columns to the source workbook columns.
     */
    public XUpdater(final XSSFWorkbook workbook, final XSSFWorkbook[] workbooks,
                    final int targetIndexCol, final int sourceIndexCol, @NotNull final HashMap<Integer, Integer> map) {
        this.target = workbook;
        this.sources = workbooks;
        this.sourcesLen = workbooks.length;
        this.targetIndexCol = targetIndexCol;
        this.sourceIndexCol = sourceIndexCol;
        this.map = map;
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

        initializeIndices();

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

    /**
     * Create index for the target workbook and a list of indices for each of the source workbooks.
     */
    private void initializeIndices() throws Exception {
        sourcesIndex = new ArrayList<>();
        targetIndex = index(target, targetIndexCol);
        for (int i = 0; i < sourcesLen; i++) {
            sourcesIndex.add(index(sources[i], sourceIndexCol));
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

    /**
     * Updates the {@link #target} with data stored in the {@link #sources} using the mapping {@link #map} between their columns.
     */
    public void update() {
        updateDuplicates();
        updatesMissing();
        updateExtra();
    }


    private void updateDuplicates() {
        for (String key : duplicates.keySet()) {
            int targetRowNum = targetIndex.get(key);
            Row targetRow = target.getSheetAt(0).getRow(targetRowNum);
            int sourceNum = duplicates.get(key);
            int sourceRowNum = sourcesIndex.get(sourceNum).get(key);
            Row sourceRow = sources[sourceNum].getSheetAt(0).getRow(sourceRowNum);
            String targetKey = targetRow.getCell(targetIndexCol).getStringCellValue();
            String sourceKey = sourceRow.getCell(sourceIndexCol).getStringCellValue();
            // cross-check control
            if (key.equals(sourceKey) && key.equals(targetKey)) {
                updateRow(targetRow, sourceRow, map);
            } else {
                System.out.println("mismatch in updating the keys! Duplicates contains: " + key + ", targetKey: " + targetKey + ", sourceKey: " + sourceKey);
            }
        }

    }

    /**
     * Adds a string cell at the end of the row which key is not present in any of the source files.
     */
    private void updatesMissing() {
        for (String key : missing) {
            int rowNum = targetIndex.get(key);
            Row row = target.getSheetAt(0).getRow(rowNum);
            Cell cell = row.createCell(row.getPhysicalNumberOfCells() + 1, Cell.CELL_TYPE_STRING);
            cell.setCellValue("Assente");
            final CellStyle style = target.createCellStyle();
            final Font font = target.createFont();
            font.setColor(HSSFColor.RED.index);
            style.setFont(font);
            cell.setCellStyle(style);
        }

    }

    private void updateExtra() {
        // TODO
    }

    /**
     * Updates targetRow with information from the sourceRow using given map as a correspondence between the row cells.
     *
     * @param targetRow
     * @param sourceRow
     */
    private void updateRow(final Row targetRow, final Row sourceRow, final HashMap<Integer, Integer> map) {
        for (int targetIndex : map.keySet()) {
            int sourceIndex = map.get(targetIndex);
            Cell sourceCell = sourceRow.getCell(sourceIndex);
            if (sourceCell == null) {
                System.out.println("source column " + sourceIndex + " is not present. Skipping it.");
                continue;
            }
            int sourceCellType = sourceCell.getCellType();
            Cell targetCell = targetRow.getCell(targetIndex);

            if (targetCell == null) {
                targetCell = targetRow.createCell(targetIndex, sourceCellType);
            }
            if (sourceCellType != targetCell.getCellType()) {
                System.out.println("cell type mismatch: " + sourceCell.getCellType() + " vs " + targetCell.getCellType()
                        + " for key " + targetRow.getCell(targetIndexCol).getStringCellValue() + ". Skipping it.");
                continue;
            }
            switch (sourceCellType) {
                case Cell.CELL_TYPE_BLANK:
                    System.out.println("source cell is blank");
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                default:
                    System.out.println("Cell type " + sourceCellType + " is not supported. Skipping the update of this cell.");
            }
        }

    }

}
