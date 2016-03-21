package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class Main {

    public static void main(String[] args) {
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
        String[] smallFiles = new String[]{"Spalm Srl.xlsx", "KGP.xlsx", "Din.xlsx"};
        String bigFile = "A008 H lavoro Riparti da Qui NON Tagliato.xlsx";
        XFileReader fr = new XFileReader(folderName + bigFile, 1);
        HashMap<String, Row> mapA = fr.loadFromFile();
        System.out.println("mapA size = " + mapA.size());

        HashMap<String, Row> mapB = new HashMap<>();
        for (String fileName : smallFiles) {
            fr = new XFileReader(folderName + fileName, 0);
            HashMap<String, Row> smallMap = fr.loadFromFile();
            smallMap.remove("Dominio");
            addCellData(smallMap, fileName);
            try {
                join(mapB, smallMap);
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
        System.out.println("mapB size = " + mapB.size());
        int common = 0, distinct = 0;

        for (String index : mapA.keySet()) {
            if (mapB.containsKey(index)) {
                common++;
                mapB.remove(index);
                // the index is present in both maps
                update(mapA, mapB.get(index), index);
            } else {
                distinct++;
            }
        }
        distinct = distinct + mapB.size();
        System.out.println(String.valueOf(common) + " keys are common");
        System.out.println(String.valueOf(distinct) + " keys are distinct");
    }

    /**
     * Updates information stored in target under key index with data in given row.
     * @param target
     * @param row
     * @param index
     */
    private static void update(HashMap<String, Row> target, Row row, String index) {


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
