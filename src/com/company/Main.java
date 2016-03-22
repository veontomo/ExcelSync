package com.company;

import com.oracle.deploy.update.Updater;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class Main {

    public static void main(String[] args) throws Exception {
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        //This data needs to be written (Object[])
//        Map<String, Object[]> data = new TreeMap<String, Object[]>();
//        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
//        data.put("2", new Object[] {1, "Amit", "Shukla"});
//        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
//        data.put("4", new Object[] {3, "John", "Adwards"});
//        data.put("5", new Object[] {4, "Brian", "Schultz"});
//
//        //Iterate over data and write to sheet
//        Set<String> keyset = data.keySet();
//        int rownum = 0;
//        for (String key : keyset)
//        {
//            Row row = sheet.createRow(rownum++);
//            Object [] objArr = data.get(key);
//            int cellnum = 0;
//            for (Object obj : objArr)
//            {
//                Cell cell = row.createCell(cellnum++);
//                if(obj instanceof String)
//                    cell.setCellValue((String)obj);
//                else if(obj instanceof Integer)
//                    cell.setCellValue((Integer)obj);
//            }
//        }
//        try
//        {
//            //Write the workbook in file system
//            FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
//            workbook.write(out);
//            out.close();
//            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
//        }
//        catch (Exception e)
//        {
//            e.printStackTrace();
//        }


//        String[] filePaths = new String[]{"excel_data\\Din.xlsx", "excel_data\\KGP.xlsx", "excel_data\\Spalm Srl.xlsx"};
//        XSSFWorkbook[] workbooks = new XSSFWorkbook[]{};
//        try
//        {
//            FileInputStream file = new FileInputStream(new File("excel_data\\A008 H lavoro Riparti da Qui NON Tagliato.xlsx"));
//
//            //Create Workbook instance holding reference to .xlsx file
//            workbook = new XSSFWorkbook(file);
//
//            //Get first/desired sheet from the workbook
//            sheet = workbook.getSheetAt(0);
//
//            //Iterate through each rows one by one
//            Iterator<Row> rowIterator = sheet.iterator();
//            while (rowIterator.hasNext())
//            {
//                Row row = rowIterator.next();
//                //For each row, iterate through all the columns
//                Iterator<Cell> cellIterator = row.cellIterator();
//
//                while (cellIterator.hasNext())
//                {
//                    Cell cell = cellIterator.next();
//                    //Check the cell type and format accordingly
//                    switch (cell.getCellType())
//                    {
//                        case Cell.CELL_TYPE_NUMERIC:
//                            System.out.print(cell.getNumericCellValue() + " ");
//                            break;
//                        case Cell.CELL_TYPE_STRING:
//                            System.out.print(cell.getStringCellValue() + " ");
//                            break;
//                    }
//                }
//                System.out.println("");
//            }
//            file.close();
//        }
//        catch (Exception e)
//        {
//            e.printStackTrace();
//        }


        String folderName = "excel_data\\";
        // the target file and list of the source files
        String target = "A008 H lavoro Riparti da Qui NON Tagliato.xlsx";
        String[] sources = new String[]{"Spalm Srl.xlsx", "KGP.xlsx", "Din.xlsx"};
        // list of strings to identify the sources
        String[] marker = new String[]{"SPALM SRL", "KGP", "DIN"};
        int sourcesLen = sources.length;
        XFileReader fr = new XFileReader();
        XSSFWorkbook workbookA = fr.loadFromFile(folderName + target);
        HashMap<String, Integer> mapA = fr.index(workbookA, 1);
        XSSFWorkbook[] workbooks = new XSSFWorkbook[sourcesLen];
        List<HashMap<String, Integer>> maps = new ArrayList<>();

        for (int i = 0; i < sourcesLen; i++) {
            workbooks[i] = fr.loadFromFile(folderName + sources[i]);
            maps.add(fr.index(workbooks[i], 0));
        }

        XUpdater updater = new XUpdater(mapA, maps);
        updater.analyze();
        HashMap<String, Integer> duplicates = updater.getDuplicates();
        HashMap<String, Integer> extra = updater.getExtra();
        List<String> missing = updater.getMissing();
        System.out.println("duplicates: " + duplicates.size());
        System.out.println("missing: " + missing.size());
        System.out.println("extra: " + extra.size());


        // first pass


//        FileOutputStream out = new FileOutputStream(new File("test.xlsx"));
//        workbookA.write(out);


//        HashMap<String, Row> mapA = fr.loadFromFile();
//        System.out.println("mapA size = " + mapA.size());
//
//        HashMap<String, Row> mapB = new HashMap<>();
//        for (int i = 0; i < smallFiles.length; i++) {
//            String fileName = smallFiles[i];
//            fr = new XFileReader(folderName + fileName, 0);
//            HashMap<String, Row> smallMap = fr.loadFromFile();
//            smallMap.remove("Dominio");
//            addCellData(smallMap, marker[i]);
//            try {
//                join(mapB, smallMap);
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
//
//        }
//        System.out.println("mapB size = " + mapB.size());
//
//        final HashMap<Integer, Integer> mapping = new HashMap<>();
//        mapping.put(5, 3);
//        mapping.put(6, 2);
//        mapping.put(7, 2);
//        mapping.put(9, 5);
//        mapping.put(10, 6);
//        mapping.put(11, 7);
//        mapping.put(12, 8);
//        mapping.put(18, 9);
//        mapping.put(22, 1);
//        int common = 0, distinct = 0;
//
//        final CellStyle style = workbook.createCellStyle();
//        final Font font = workbook.createFont();
//        font.setColor(HSSFColor.RED.index);
//        style.setFont(font);
//
//        int cellsNumA = mapA.get(mapA.keySet().iterator().next()).getPhysicalNumberOfCells();
//        // first pass: iterate ove keys in mapA and remove those keys if they are present in mapB
//        for (String index : mapA.keySet()) {
//            if (mapB.containsKey(index)) {
//                common++;
//                // the index is present in both maps
//                try {
//                    update(mapA.get(index), mapB.get(index), index, mapping);
//                } catch (Exception e) {
//                    e.printStackTrace();
//                    System.out.println("failed to update row corresponding to " + index + ", error: " + e.getMessage());
//                }
//                mapB.remove(index);
//            } else {
//                distinct++;
//                addCell(mapA.get(index), cellsNumA, "Assente", style);
//            }
//        }
//
//        // second pass: iterate over remaining keys in mapB
//        for (String index : mapB.keySet()) {
//            Row r = sheet.createRow(cellsNumA + 1);
//            populateRow(r, mapB.get(index), mapping);
//            Cell c = r.createCell(r.getPhysicalNumberOfCells(), Cell.CELL_TYPE_STRING);
//            c.setCellValue("Nuovo");
//            mapA.put(index, r);
//        }
//        distinct = distinct + mapB.size();
//        System.out.println(String.valueOf(common) + " keys are common");
//        System.out.println(String.valueOf(distinct) + " keys are distinct");
//
//        save(mapA, folderName + "result.xlsx");

    }

    /**
     * Saves data in the file
     *
     * @param mapA
     * @param s
     */
    private static void save(HashMap<String, Row> mapA, String s) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("test");
        int rownum = 0;
        for (String key : mapA.keySet()) {
            Row row = sheet.createRow(rownum++);

            try {
                copy(row, mapA.get(key));
            } catch (Exception e) {
                System.out.println("Exception: key " + key + ", rownum " + rownum + ", message: " + e.getMessage());
//                return;
            }
        }
        try {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(s));
            workbook.write(out);
            out.close();
            System.out.println(s + " written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * Copies the source row into target one.
     *
     * @param target
     * @param source
     */
    private static void copy(Row target, Row source) throws Exception {
        Iterator<Cell> iterator = source.iterator();
        int cellCounter = 0;
        while (iterator.hasNext()) {
            Cell sourceCell = iterator.next();
            int type = sourceCell.getCellType();
            Cell targetCell = target.createCell(cellCounter, type);
            switch (type) {
                case Cell.CELL_TYPE_NUMERIC:
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                default:
                    throw new Exception("Unsupported cell type " + type + ", cellCounter " + cellCounter);
            }
//            CellStyle newStyle = workbook.createCellStyle();
//            newStyle.cloneStyleFrom(sourceCell.getCellStyle());
//            targetCell.setCellStyle(newStyle);
            cellCounter++;


        }


    }

    /**
     * Populates a target row with the source row using given mapping between their cells
     *
     * @param target
     * @param source
     * @param mapping
     */
    private static void populateRow(Row target, Row source, HashMap<Integer, Integer> mapping) throws Exception {
        for (int targetCellNum : mapping.keySet()) {
            int sourceCellNum = mapping.get(targetCellNum);
            Cell sourceCell = source.getCell(sourceCellNum);
            Cell targetCell = target.getCell(targetCellNum);
            if (targetCell == null) {
                targetCell = target.createCell(targetCellNum);
                targetCell.setCellType(sourceCell.getCellType());
            }
            updateCell(targetCell, sourceCell);
        }
    }

    /**
     * Updates target cell  with the data from source cell.
     *
     * @param targetCell
     * @param sourceCell
     */
    private static void updateCell(Cell targetCell, Cell sourceCell) throws Exception {
        int cellType = sourceCell.getCellType();
        targetCell.setCellType(cellType);
        switch (cellType) {
            case Cell.CELL_TYPE_STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            default:
                throw new Exception("Unsupported cell type: " + cellType);
        }

    }


    /**
     * Adds a cell at the end of the row with given string content and apply given style.
     *
     * @param row
     * @param marker
     * @param style
     */
    private static void addCell(Row row, int pos, String marker, CellStyle style) {
        Cell c = row.createCell(pos, Cell.CELL_TYPE_STRING);
        c.setCellValue(marker);
//        c.setCellStyle(style);

    }

    /**
     * Updates information stored in target row under key index with data in given row.
     *
     * @param target
     * @param info
     * @param index
     * @param mapping correspondence between row cells of the target and info
     */
    private static void update(Row target, final Row info, final String index, final HashMap<Integer, Integer> mapping) throws Exception {
        for (int pos : mapping.keySet()) {
            Cell targetCell = target.getCell(pos);
            Cell infoCell = info.getCell(mapping.get(pos));
            if (targetCell == null) {
                targetCell = target.createCell(pos, infoCell.getCellType());
                System.out.println("created cell for pos " + pos);
            }
            int targetCellType = targetCell.getCellType();
            int infoCellType = infoCell.getCellType();

            if (infoCellType != targetCellType) {
                System.out.println("cell type mismatch: " + targetCellType + " vs " + infoCellType + " for " + index + ", pos = " + pos + ", mapping: " + mapping.get(pos));
                System.out.println("Imposing type " + infoCellType);
                targetCell.setCellType(infoCellType);
            }

            if (infoCellType == Cell.CELL_TYPE_NUMERIC) {
                targetCell.setCellValue(infoCell.getNumericCellValue());
            }

            switch (infoCellType) {
                case Cell.CELL_TYPE_STRING:
                    targetCell.setCellValue(infoCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    targetCell.setCellValue(infoCell.getNumericCellValue());
                    break;
                default:
                    throw new Exception("Unsupported cell type: " + infoCellType + " for " + index + ", pos = " + pos + ", mapping: " + mapping.get(pos));


            }
        }

    }


    /**
     * Joins two hash maps.
     * If there is an index in common, an exception will be thrown.
     *
     * @param big
     * @param data
     * @return
     */
    private static void join(HashMap<String, Row> big, final HashMap<String, Row> data) throws Exception {
        for (String index : data.keySet()) {
            if (big.containsKey(index)) {
                throw new Exception("key " + index + " is already present!");
            }
            big.put(index, data.get(index));
        }
    }

    /**
     * Modifies a hash map by adding a cell to the end of each row with given string content.
     *
     * @param hashMap a hash map
     * @param marker  a content of the cell
     * @return the reference
     */
    private static void addCellData(HashMap<String, Row> hashMap, String marker) {
        for (String item : hashMap.keySet()) {
            Row row = hashMap.get(item);
            Cell cell = row.createCell(row.getPhysicalNumberOfCells(), Cell.CELL_TYPE_STRING);
            cell.setCellValue(marker);
        }
    }
}
