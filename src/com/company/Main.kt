package com.company

import org.apache.commons.cli.DefaultParser
import org.apache.commons.cli.Option
import org.apache.commons.cli.Options
import java.io.File
import java.io.FileOutputStream


fun main(args: Array<String>) {
    val TOKEN_DIR = "d"
    val TOKEN_TARGET = "t"
    val TOKEN_SOURCES = "s"
    val TOKEN_OUT = "o"
    val options = Options()
    options.addOption(TOKEN_DIR, true, "set the working folder")
    options.addOption(TOKEN_TARGET, true, "set the target file name in the working folder")
    options.addOption(TOKEN_OUT, true, "set the output file name to be saved in the working folder")
    val option = Option(TOKEN_SOURCES, "set the source file names and their aliases")
    option.args = Option.UNLIMITED_VALUES
    options.addOption(option)

    val parser = DefaultParser()
    val cmd = parser.parse(options, args)
    if (!cmd.hasOption(TOKEN_DIR)) {
        println("No working directory is set")
        return
    }

    val folderName = cmd.getOptionValue(TOKEN_DIR)
    println("working folder: $folderName")

    if (!cmd.hasOption(TOKEN_TARGET)) {
        println("No target file is set.")
        return
    }
    val target = folderName + cmd.getOptionValue(TOKEN_TARGET)
    println("path to the target file: $target")

    if (!cmd.hasOption(TOKEN_SOURCES)) {
        println("No source files are set.")
        return
    }

    val sourcesRaw = cmd.getOptionValues(TOKEN_SOURCES)
    if (!cmd.hasOption(TOKEN_OUT)) {
        println("No output file is set.")
        return
    }
    val outfile = folderName + cmd.getOptionValue(TOKEN_OUT)
    println("output file name: $outfile")

    val len = sourcesRaw.size
    if (len % 2 != 0) {
        println("Each file name must be preceded by its alias, instead the following is given: ${sourcesRaw.joinToString { it }}")
        return
    }

    val sources = mutableMapOf<String, String>()
    for (i in 0..(len - 2) step 2) {
        sources.put(sourcesRaw[i], folderName + sourcesRaw[i + 1])
    }
    println("source files: ${sources.map { "${it.value} as ${it.key}"}.joinToString { it }}")


    val fr = XFileReader()
    val workbookA = fr.loadFromFile(target)
    val workbooks = sources.map { it.key to fr.loadFromFile(it.value) }.toMap()


    // correspondence between columns of the target workbook and the source workbooks.
    val mapping = mapOf(1 to 0, 5 to 3, 6 to 3, 7 to 2, 8 to 5, 9 to 6, 10 to 7, 11 to 8, 12 to 9, 15 to 4)

    // Strings to be added at the end of the updated rows
    val markers = arrayOf("Aggiornato", "Nuovo", "Assente")
    // List of strings to be ignored when creating the index of each workbook
    val blacklist = listOf("Dominio", "Descrizione Sito")

    val updater = XUpdater(workbookA, workbooks, 1, 0, mapping, markers, blacklist)

    updater.analyze()

    val duplicates = updater.duplicates
    val extra = updater.extra
    val missing = updater.missing
    println("duplicates: ${duplicates.size} items:\n ${duplicates.map { it.key + " -> " + it.value }.joinToString(", ", "", "", 5, "...", { it })} ")
    println("missing: ${missing.size} item:\n ${missing.joinToString(", ", "", "", 5, "...", { it })}")
    println("extra: ${extra.size} items:\n ${extra.map { it.key + " -> " + it.value }.joinToString(", ", "", "", 5, "...", { it })} ")

//    updater.update()

    val out = FileOutputStream(File(outfile))
    workbookA.write(out)
}


