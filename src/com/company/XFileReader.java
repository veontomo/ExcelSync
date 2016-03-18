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
     * Loads data from a given excel file
     * @param path
     * @return
     */
    public HashMap<String, Row> loadFromFile(String path){
        HashMap<String, Row> data = new HashMap<>();
        try
        {
            FileInputStream file = new FileInputStream(new File(path));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            String key;
            Row row;
            while (rowIterator.hasNext())
            {
                row = rowIterator.next();
                key = row.getCell(1).getStringCellValue();

                if (data.containsKey(key)){
                    System.out.println("duplicate key: " + key);
                } else {
                    data.put(key, row);
                }

                //For each row, iterate through all the columns
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
            }
            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return data;
    }
}
