package com.company;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class Main {

    public static void main(String[] args) throws Exception {

        String folderName = "excel_data\\";
        // the target file and list of the source files
        String target = "A008 H lavoro Riparti da Qui NON Tagliato.xlsx";
        String[] sources = new String[]{"Spalm_Srl_with_filename.xlsx", "KGP_with_filename.xlsx", "Din_with_filename.xlsx"};

        int sourcesLen = sources.length;
        XFileReader fr = new XFileReader();
        XSSFWorkbook workbookA = fr.loadFromFile(folderName + target);
        XSSFWorkbook[] workbooks = new XSSFWorkbook[sourcesLen];

        for (int i = 0; i < sourcesLen; i++) {
            workbooks[i] = fr.loadFromFile(folderName + sources[i]);

        }
        /**
         * correspondence between columns of the target workbook and the source workbooks.
         */
        final HashMap<Integer, Integer> mapping = new HashMap<>();
        mapping.put(5, 3);
        mapping.put(6, 4);
        mapping.put(7, 2);
        mapping.put(9, 5);
        mapping.put(10, 6);
        mapping.put(11, 7);
        mapping.put(12, 8);
        mapping.put(18, 9);
        mapping.put(22, 1);

        /**
         * Strings to be added at the and of the updated rows
         */
        final String[] markers = new String[]{"Aggiornato", "Nuovo", "Assente"};
        /**
         * List of strings to be ignored when creating the index of each workbook
         */
        final List<String> blacklist = new ArrayList<>();
        blacklist.add("Dominio");
        blacklist.add("Descrizione Sito");

        XUpdater updater = new XUpdater(workbookA, workbooks, 1, 0, mapping, markers, blacklist);

        updater.analyze();

        HashMap<String, Integer> duplicates = updater.getDuplicates();
        HashMap<String, Integer> extra = updater.getExtra();
        List<String> missing = updater.getMissing();
        System.out.println("duplicates: " + duplicates.size());
        System.out.println("missing: " + missing.size());
        System.out.println("extra: " + extra.size());

        updater.update();

        FileOutputStream out = new FileOutputStream(new File(folderName + "updated.xlsx"));
        workbookA.write(out);
    }
}
