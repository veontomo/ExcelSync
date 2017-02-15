package com.company;

import com.sun.istack.internal.NotNull;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Updates a workbook with data from other workbooks
 */
public class XUpdater {

    private final Map<String, XSSFWorkbook> sources;
    private final int sourcesLen;
    private final XSSFWorkbook target;
    private final Map<Integer, Integer> map;

    private final int targetIndexCol;
    private final int sourceIndexCol;
    private final List<String> blacklist;

    private HashMap<String, Integer> targetIndex;
    /**
     * Collection of indexes of all source files. The structure is as follows:
     * [alias_1 => index_1, alias_2 => index_2, ...]
     * where each index has the following structure:
     * [key_1 => pos_1, key_2 => pos_2, ...]
     * where pos_i stands for the number of row containing key_i.
     */
    private Map<String, Map<String, Integer>> sourcesIndex;


    /**
     * list of keys that are present in both targetIndex and sourcesIndex
     */

    private HashMap<String, String> duplicates;
    /**
     * list of keys that are present in targetIndex and not present in any of the sourcesIndex
     */
    private List<String> missing;

    /**
     * Contains strings that are present in one of the source indexes and NOT present in the target index.
     * The values of this map are aliases of the source files in which the above mentioned strings are found.
     */
    private HashMap<String, String> extra;

    /**
     * Style to be applied to a cell that is to be appended to the rows present in {@link #missing}
     */
    private final CellStyle styleForMissing;

    /**
     * Style to be applied to a cell that is to be appended to the rows present in {@link #duplicates}
     */
    private final CellStyle styleForDuplicates;

    /**
     * Style to be applied to a cell that is to be appended to the rows present in {@link #extra}
     */
    private final CellStyle styleForExtra;
    private final String markerForDuplicates;
    private final String markerForExtra;
    private final String markerForMissing;

    private final CellStyle dateCellStyle;

    /**
     * Constructor.
     *
     * @param workbook       a target workbook
     * @param workbooks      array of source workbooks
     * @param targetIndexCol a number of the column of the target workbook w.r.t. which an index is to be constructed
     * @param sourceIndexCol a number of the column of the source workbooks w.r.t. which an index is to be constructed
     * @param map            defines the mapping from the target workbook columns to the source workbook columns.
     */
    public XUpdater(final XSSFWorkbook workbook, final Map<String, XSSFWorkbook> workbooks,
                    final int targetIndexCol, final int sourceIndexCol, @NotNull final Map<Integer, Integer> map,
                    final String[] markers, final List<String> blacklist) {
        this.target = workbook;
        this.sources = workbooks;
        this.sourcesLen = workbooks.size();
        this.targetIndexCol = targetIndexCol;
        this.sourceIndexCol = sourceIndexCol;
        this.map = map;
        this.blacklist = blacklist;

        this.styleForMissing = target.createCellStyle();
        final Font font = target.createFont();
        font.setColor(HSSFColor.RED.index);
        styleForMissing.setFont(font);

        this.styleForDuplicates = target.createCellStyle();
        final Font font2 = target.createFont();
        font2.setColor(HSSFColor.BLUE.index);
        styleForDuplicates.setFont(font2);

        this.styleForExtra = target.createCellStyle();
        final Font font3 = target.createFont();
        font3.setColor(HSSFColor.GREEN.index);
        styleForExtra.setFont(font3);

        this.markerForDuplicates = markers[0];
        this.markerForExtra = markers[1];
        this.markerForMissing = markers[2];

        // date formatter
        dateCellStyle = target.createCellStyle();
        CreationHelper createHelper = target.getCreationHelper();
        dateCellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("m/d/yy"));

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
            for (String alias : sourcesIndex.keySet()) {

                if (sourcesIndex.get(alias).containsKey(key)) {
                    if (duplicates.containsKey(key)) {
                        throw new Exception("key " + key + " has already been found in source with alias " + alias + ". Resolve to proceed.");
                    }
                    isFoundInSources = true;
                    duplicates.put(key, alias);
                }
            }
            if (!isFoundInSources) {
                missing.add(key);
            }
        }
        // second pass: iterate over the sourcesIndex and control if they contain keys that are not in the targetIndex
        for (String alias : sources.keySet()) {
            for (String key : sourcesIndex.get(alias).keySet()) {
                if (targetIndex.containsKey(key)) {
                    // cross check: the variable "duplicates" must contain this key as well.
                    if (duplicates.containsKey(key) && duplicates.get(key).equals(alias)) {
                        System.out.println("cross-check is OK");
                    } else {
                        System.out.println("cross-check is not OK for key " + key + " that is supposed to be in set " + alias);
                    }
                } else {
                    if (extra.containsKey(key)) {
                        System.out.println("key " + key + " is found in source n. " + alias + ", while it has already been added to the extra index.");
                    } else {
                        extra.put(key, alias);
                    }
                }
            }
        }
    }

    /**
     * Create index for the target workbook and a list of indices for each of the source workbooks.
     */
    private void initializeIndices() throws Exception {
        sourcesIndex = new HashMap<>();
        targetIndex = index(target, targetIndexCol);
        for (String key : sources.keySet()) {
            sourcesIndex.put(key, index(sources.get(key), sourceIndexCol));

        }
    }

    public Map<String, String> getDuplicates() {
        return duplicates;
    }

    public List<String> getMissing() {
        return missing;
    }

    public HashMap<String, String> getExtra() {
        return extra;
    }

    /**
     * Create an index of given workbook: a map from string content of cells of given column to the number of row
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
            if (blacklist.contains(key)) {
                System.out.println("Key \"" + key + "\" is listed in the blacklist and hence is not added to the index.");
                continue;
            }
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
            String alias = duplicates.get(key);
            int sourceRowNum = sourcesIndex.get(alias).get(key);
            Row sourceRow = sources.get(alias).getSheetAt(0).getRow(sourceRowNum);
            String targetKey = targetRow.getCell(targetIndexCol).getStringCellValue();
            String sourceKey = sourceRow.getCell(sourceIndexCol).getStringCellValue();
//             cross-check control
            if (key.equals(sourceKey) && key.equals(targetKey)) {
                updateRow(targetRow, sourceRow, map);
            } else {
                System.out.println("mismatch in updating the keys! Duplicates contains: " + key + ", targetKey: " + targetKey + ", sourceKey: " + sourceKey);
            }
            markRow(targetRow, 25, markerForDuplicates, styleForDuplicates);
        }

    }

    /**
     * Adds a string cell at the end of the row which key is not present in any of the source files.
     */
    private void updatesMissing() {
        HashMap<Integer, Integer> map = new HashMap<>();
        for (String key : missing) {
            int rowNum = targetIndex.get(key);
            Row row = target.getSheetAt(0).getRow(rowNum);
            updateRow(row, null, map);
            markRow(row, 25, markerForMissing, styleForMissing);
        }


    }

    private void updateExtra() {
        for (String key : extra.keySet()) {
            String alias = extra.get(key);
            int sourceRowNum = sourcesIndex.get(alias).get(key);
            Row sourceRow = sources.get(alias).getSheetAt(0).getRow(sourceRowNum);
            int totalRowNum = target.getSheetAt(0).getLastRowNum();
            Row targetRow = target.getSheetAt(0).createRow(totalRowNum + 1);
            targetRow.createCell(targetIndexCol, Cell.CELL_TYPE_STRING).setCellValue(key);
            updateRow(targetRow, sourceRow, map);
            // set up by hand
            Map<Integer, String> data = new HashMap<>();
            data.put(2, "Confermato");
            data.put(3, "Confermato");
            data.put(13, "Sì");

            data.put(16, "Dominiando");
            data.put(17, alias);
            data.put(18, alias + " SRL");
            data.put(19, key);
            data.put(23, key);

            try {
                fillInRowCells(targetRow, data);
            } catch (Exception e) {
                System.out.println("Error when adjusting an extra row for " + key + " from " + alias + ": " + e.getMessage());
            }
            targetRow.getCell(5).setCellStyle(dateCellStyle);
            targetRow.getCell(6).setCellStyle(dateCellStyle);
            targetRow.getCell(7).setCellStyle(dateCellStyle);
            Cell cell = targetRow.createCell(22);
            cell.setCellValue(targetRow.getCell(7).getDateCellValue());
            cell.setCellStyle(dateCellStyle);

            markRow(targetRow, 25, markerForExtra, styleForExtra);

        }
    }

    /**
     * Create cells in the row and fill them in with given strings.
     *
     * @param row  the row whose cell are to be filled in
     * @param data map from cell numbers to string that the cell should contain.
     * @throws Exception if the row already contains at least one cell that should be filled in.
     */
    private void fillInRowCells(Row row, Map<Integer, String> data) throws Exception {
        for (Integer index : data.keySet()) {
            Cell cell = row.getCell(index);
            if (cell == null) {
                cell = row.createCell(index, Cell.CELL_TYPE_STRING);
            } else {
                throw new Exception("Cell n. " + index + " already exists! It contains: " + cell.getStringCellValue());
            }
            cell.setCellValue(data.get(index));
        }

    }

    /**
     * Updates targetRow with information from the sourceRow using given map as a correspondence between the row cells.
     *
     * @param targetRow
     * @param sourceRow
     * @param map
     */
    private void updateRow(final Row targetRow, final Row sourceRow, final Map<Integer, Integer> map) {
        for (int targetCellNum : map.keySet()) {
            int sourceCellNum = map.get(targetCellNum);
            Cell sourceCell = sourceRow.getCell(sourceCellNum);
            if (sourceCell == null) {
                System.out.println("source column " + sourceCellNum + " is not present. Skipping it.");
                continue;
            }
            int sourceCellType = sourceCell.getCellType();
            Cell targetCell = targetRow.getCell(targetCellNum);

            if (targetCell != null && sourceCellType != targetCell.getCellType()) {
                System.out.println("cell type mismatch: " + sourceCell.getCellType() + " vs " + targetCell.getCellType()
                        + " for key " + targetRow.getCell(targetIndexCol).getStringCellValue() + ". Skipping it.");
                continue;
            }
            if (targetCell == null) {
                targetCell = targetRow.createCell(targetCellNum, sourceCellType);
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


    /**
     * Insert a marker at a cell with given number and apply cell styles.
     *
     * @param targetRow
     * @param marker
     * @param style
     * @param cellNum   cell number (zero based) at which to insert the marker. -1 in order to insert at the end of the row.
     */
    private void markRow(Row targetRow, int cellNum, String marker, CellStyle style) {
        int pos = cellNum == -1 ? targetRow.getLastCellNum() + 1 : cellNum;
        Cell cell = targetRow.getCell(cellNum);
        if (cell == null) {
            cell = targetRow.createCell(pos, Cell.CELL_TYPE_STRING);
        }
        if (marker != null) {
            cell.setCellValue(marker);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

}
