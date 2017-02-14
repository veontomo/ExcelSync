package com.company

import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.io.File
import java.io.FileOutputStream
import java.sql.*
import java.util.*


fun main(args: Array<String>) {

    if (args.size < 2){
        println("At leas three arguments are expected.")
    }

    val folderName = args[0]
    // the target file and list of the source files
    val target = folderName + args[1]
    val sources = Arrays.copyOfRange(args, 2, args.size).map{ it -> folderName + it}


    println("target file: $target")
    println("source files: ${sources.joinToString { it }}")
    val sourcesLen = sources.size
    val fr = XFileReader()
    val workbookA = fr.loadFromFile(target)
    val workbooks = arrayOfNulls<XSSFWorkbook>(sourcesLen)

    for (i in 0..sourcesLen - 1) {
        workbooks[i] = fr.loadFromFile(sources[i])

    }
    // correspondence between columns of the target workbook and the source workbooks.
    val mapping = mapOf(2 to 1, 6 to 4, 7 to 4, 8 to 3, 9 to 6, 10 to 7, 11 to 8, 12 to 9, 13 to 19, 16 to 5)

    // Strings to be added at the end of the updated rows
    val markers = arrayOf("Aggiornato", "Nuovo", "Assente")
    // List of strings to be ignored when creating the index of each workbook
    val blacklist = ArrayList<String>()
    blacklist.add("Dominio")
    blacklist.add("Descrizione Sito")

    val updater = XUpdater(workbookA, workbooks, 1, 0, mapping, markers, blacklist)

    updater.analyze()

    val duplicates = updater.duplicates
    val extra = updater.extra
    val missing = updater.missing
    println("duplicates: " + duplicates.size)
    println("missing: " + missing.size)
    println("extra: " + extra.size)

    //        updater.update();

    //        FileOutputStream out = new FileOutputStream(new File(folderName + "updated.xlsx"));
    //        workbookA.write(out);
}

fun dbRead(dbName: String, tblName: String) {
    val connectionProps = Properties()
    connectionProps.put("user", "siti_local")
    connectionProps.put("password", "siti_local_read")
    val pattern = "[^\\p{Alnum}_]"
    val dbName2 = dbName.replace(pattern.toRegex(), "")
    val tblName2 = tblName.replace(pattern.toRegex(), "")
    try {
        val conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/" + dbName2, connectionProps)
        val statement = conn.createStatement()
        val result = statement.executeQuery("SELECT * FROM $tblName2;")
        print(result.fetchSize)
        while (result.next()) {
            println(result.getString(2))
        }


        conn.close()
    } catch (e: SQLException) {
        e.printStackTrace()
    }

}

