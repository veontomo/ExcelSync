package com.company

import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io.File
import java.io.FileInputStream

/**
 * Performs operations with excel files.
 */
class XFileReader {
    /**
     * Read an Excel file and return it as Workbook instance.

     * @param filePath a path to the file to read from
     * *
     * @return a XSSWorkbook instance
     */
    fun loadFromFile(filePath: String): XSSFWorkbook {
        try {
            val f = File(filePath)
            val file = FileInputStream(f)
            val workbook = XSSFWorkbook(file)
            println("Loaded " + workbook.getSheetAt(0).physicalNumberOfRows + " rows from file " + filePath)
            file.close()
            return workbook
        } catch (e: Exception) {
            println("Error " + e.message + " when processing file " + filePath)
            println("Try to open the file by Excel and save it with extension .xlsx")
        }

    }


}
