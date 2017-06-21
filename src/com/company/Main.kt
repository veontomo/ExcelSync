package com.company

import org.apache.commons.cli.*
import java.io.File
import java.io.FileOutputStream


fun main(args: Array<String>) {
    val TOKEN_DIR = "d"
    val TOKEN_TARGET = "t"
    val TOKEN_SOURCES = "s"
    val TOKEN_OUT = "o"
    println(args.joinToString { it })
    val options = Options()

    val optionWorkDir = Option.builder(TOKEN_DIR).argName("dir").desc("set the working folder").hasArg().required().build()

    val optionTargetFile = Option.builder(TOKEN_TARGET)
            .argName("file")
            .desc("set the target file name in the working folder")
            .hasArg()
            .required()
            .build()
    val optionOutputFile = Option.builder(TOKEN_OUT)
            .argName("file")
            .desc("set the output file name to be saved in the working folder")
            .hasArg()
            .required()
            .build()
    val optionSourceFiles = Option.builder(TOKEN_SOURCES)
            .argName("alias=file")
            .desc("set the source file names and their aliases")
            .numberOfArgs(2)
            .required(false)
            .valueSeparator('=')
            .build()

    options.addOption(optionWorkDir)
    options.addOption(optionTargetFile)
    options.addOption(optionOutputFile)
    options.addOption(optionSourceFiles)


    val parser = DefaultParser()
    val cmd: CommandLine = try {
        parser.parse(options, args)
    } catch (e: ParseException) {
        val formatter = HelpFormatter()
        formatter.printHelp("ExcelSync", options)
        println(e.message)
        return
    }

    if (!cmd.hasOption(TOKEN_DIR)) {
        println("Working folder is not set.")
        return
    }
    val rawFolderName = cmd.getOptionValue(TOKEN_DIR)
    val separ = File.separator
    val folderName = rawFolderName + (if (rawFolderName.endsWith(separ)) "" else separ)
    println("working folder: $folderName")

    if (!cmd.hasOption(TOKEN_TARGET)) {
        println("No target file is set.")
        return
    }
    val target = folderName + cmd.getOptionValue(TOKEN_TARGET)
    println("path to the target file: $target")

    if (!cmd.hasOption(TOKEN_OUT)) {
        println("No output file is set.")
        return
    }
    val outfile = folderName + cmd.getOptionValue(TOKEN_OUT)
    println("output file name: $outfile")

    if (!cmd.hasOption(TOKEN_SOURCES)) {
        println("No source files are set.")
        return
    }

    val sourcesRaw = cmd.getOptionValues(TOKEN_SOURCES)
    println("sources raw: ${sourcesRaw.joinToString { it }}")
    val len = sourcesRaw.size
    if (len % 2 != 0) {
        println("Each file name must be preceded by its alias, instead the following is given: ${sourcesRaw.joinToString { it }}")
        return
    }

    val sources = mutableMapOf<String, String>()
    for (i in 0..(len - 2) step 2) {
        sources.put(sourcesRaw[i], folderName + sourcesRaw[i + 1])
    }
    println("source files: ${sources.map { "${it.value} as ${it.key}" }.joinToString { it }}")
    val fr = XFileReader()
    val workbookA = fr.loadFromFile(target)
    val workbooks = sources.map { it.key to fr.loadFromFile(it.value) }.toMap()


    // correspondence between columns of the target workbook and the source workbooks.
    val mapping = mapOf(1 to 0, 5 to 3, 6 to 3, 7 to 2, 8 to 5, 9 to 6, 10 to 7, 11 to 8, 12 to 9, 15 to 4)

    // List of strings to be ignored when creating the index of each workbook
    val blacklist = listOf("Dominio", "Descrizione Sito")

    val updater = XUpdater(workbookA, workbooks, 1, 0, mapping, blacklist)

    updater.analyze()

    val duplicates = updater.duplicates
    val extra = updater.extra
    val missing = updater.missing
    println("duplicates: ${duplicates.size} items:\n ${duplicates.map { it.key + " -> " + it.value }.joinToString(", ", "", "", 5, "...", { it })} ")
    println("missing: ${missing.size} item:\n ${missing.joinToString(", ", "", "", 5, "...", { it })}")
    println("extraMap: ${extra.size} items:\n ${extra.map { it.key + " -> " + it.value }.joinToString(", ", "", "", 5, "...", { it })} ")

    updater.update()

    val out = FileOutputStream(File(outfile))
    workbookA.write(out)
}


