package com.company

import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.util.HashMap

/**
 * Updates a workbook with data from other workbooks
 */
/**

 * @param workbook       a target workbook
 * *
 * @param workbooks      array of source workbooks
 * *
 * @param targetIndexCol a number of the column of the target workbook w.r.t. which an index is to be constructed
 * *
 * @param sourceIndexCol a number of the column of the source workbooks w.r.t. which an index is to be constructed
 * *
 * @param map            defines the mapping from the target workbook columns to the source workbook columns.
 */
class XUpdater(private val target: XSSFWorkbook, private val sources: Map<String, XSSFWorkbook>,
               private val targetIndexCol: Int, private val sourceIndexCol: Int, private val map: Map<Int, Int>,
               private val blacklist: List<String>) {
    private val sourcesLen: Int

    private var targetIndex: Map<String, Int> = index(target, targetIndexCol)
    /**
     * Collection of indexes of all source files. The structure is as follows:
     * [alias_1 => index_1, alias_2 => index_2, ...]
     * where each index has the following structure:
     * [key_1 => pos_1, key_2 => pos_2, ...]
     * where pos_i stands for the number of row containing key_i.
     */
    private var sourcesIndex: Map<String, Map<String, Int>> = sources.map { it -> Pair(it.key, index(it.value, sourceIndexCol)) }.toMap()


    /**
     * list of keys that are present in both targetIndex and sourcesIndex
     */

    private var duplicatesMap: MutableMap<String, String> = mutableMapOf()

    val duplicates: Map<String, String>
        get() {
            return duplicatesMap.map { Pair(it.key, it.value) }.toMap()
        }
    /**
     * list of keys that are present in targetIndex and not present in any of the sourcesIndex
     */
    private val missingList = mutableListOf<String>()

    val missing: List<String>
        get() {
            return missingList.map { it }

        }

    private val extraMap = mutableMapOf<String, String>()
    /**
     * Contains strings that are present in one of the source indexes and NOT present in the target index.
     * The values of this map are aliases of the source files in which the above mentioned strings are found.
     */
    val extra: Map<String, String>
        get() {
            return extraMap
        }

    /**
     * Style to be applied to a cell that is to be appended to the rows present in [.missingList]
     */
    private val styleForMissing: CellStyle

    /**
     * Style to be applied to a cell that is to be appended to the rows present in [.duplicatesMap]
     */
    private val styleForDuplicates: CellStyle

    /**
     * Style to be applied to a cell that is to be appended to the rows present in [.extraMap]
     */
    private val styleForExtra: CellStyle
    private val markerForDuplicates = "Aggiornato"
    private val markerForExtra = "Nuovo"
    private val markerForMissing = "Assente"

    private val dateCellStyle: CellStyle

    init {

        targetIndex = index(target, targetIndexCol)

        this.sourcesLen = sources.size

        this.styleForMissing = target.createCellStyle()
        val font = target.createFont()
        font.color = HSSFColor.RED.index
        styleForMissing.setFont(font)

        this.styleForDuplicates = target.createCellStyle()
        val font2 = target.createFont()
        font2.color = HSSFColor.BLUE.index
        styleForDuplicates.setFont(font2)

        this.styleForExtra = target.createCellStyle()
        val font3 = target.createFont()
        font3.color = HSSFColor.GREEN.index
        styleForExtra.setFont(font3)

        // date formatter
        dateCellStyle = target.createCellStyle()
        val createHelper = target.creationHelper
        dateCellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("m/d/yy"))

    }


    /**
     * Finds the sourcesIndex that contain a key from the targetIndex.
     *
     *
     * If a key is found in multiple sourcesIndex, an exception is thrown.
     *
     *
     * Returns a hash map from a string to an integer that is the ordinal number of the source in the source list in which
     * the key is found.

     * @return
     */
    @Throws(Exception::class)
    fun analyze() {
//        duplicatesMap = HashMap<String, String>()
//        missingList = ArrayList<String>()
//        extraMap = HashMap<String, String>()

//        initializeIndices()

        var isFoundInSources: Boolean
        // first pass: iterate over the targetIndex and control the presence in the sourcesIndex
        for (key in targetIndex.keys) {
            isFoundInSources = false
            for (alias in sourcesIndex.keys) {

                if (sourcesIndex[alias]!!.containsKey(key)) {
                    if (duplicatesMap!!.containsKey(key)) {
                        throw Exception("key $key has already been found in source with alias $alias. Resolve to proceed.")
                    }
                    isFoundInSources = true
                    duplicatesMap!!.put(key, alias)
                }
            }
            if (!isFoundInSources) {
                missingList!!.add(key)
            }
        }
        // second pass: iterate over the sourcesIndex and control if they contain keys that are not in the targetIndex
        for (alias in sources.keys) {
            for (key in sourcesIndex[alias]!!.keys) {
                if (targetIndex!!.containsKey(key)) {
                    // cross check: the variable "duplicatesMap" must contain this key as well.
                    if (!(duplicatesMap!!.containsKey(key) && duplicatesMap!![key] == alias)) {
                        println("cross-check is not OK for key $key that is supposed to be in set $alias")
                    }

                } else {
                    if (extraMap!!.containsKey(key)) {
                        println("key $key is found in source n. $alias, while it has already been added to the extraMap index.")
                    } else {
                        extraMap!!.put(key, alias)
                    }
                }
            }
        }
    }

    /**
     * Create index for the target workbook and a list of indices for each of the source workbooks.
     */
    @Throws(Exception::class)
    private fun initializeIndices() {
//        sourcesIndex = HashMap<String, Map<String, Int>>()
//        targetIndex = index(target, targetIndexCol)
//        for (key in sources.keys) {
//            sourcesIndex!!.put(key, index(sources[key], sourceIndexCol))
//
//        }
    }


    /**
     * Create an index of given workbook: a map from string content of cells of given column to the number of row
     * in which that string is found.

     * @param workbook
     * *
     * @param column
     * *
     * @return
     */
    @Throws(Exception::class)
    fun index(workbook: XSSFWorkbook, column: Int): Map<String, Int> {
        val map = HashMap<String, Int>()
        val sheet = workbook.getSheetAt(0)

        val rowsNum = sheet.physicalNumberOfRows
        var row: Row
        var cell: Cell
        var key: String
        for (i in 0..rowsNum - 1) {
            row = sheet.getRow(i)
            cell = row.getCell(column)
            if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
                throw Exception("Cell $column at row $i is not of string type!")
            }
            key = cell.getStringCellValue()
            if (blacklist.contains(key)) {
                println("Key \"$key\" is listed in the blacklist and hence is not added to the index.")
                continue
            }
            if (map.containsKey(key)) {
                throw Exception("Duplicate key: " + key)
            }
            map.put(key, i)
        }
        return map
    }

    /**
     * Updates the [.target] with data stored in the [.sources] using the mapping [.map] between their columns.
     */
    fun update() {
        updateDuplicates()
        updatesMissing()
        updateExtra()
    }

    /**
     * Update rows of the target workbook that are present as well in source workbooks.
     */
    private fun updateDuplicates() {
        for (key in duplicatesMap.keys) {
            val targetRowNum = targetIndex[key]
            val targetRow = target.getSheetAt(0).getRow(targetRowNum!!)
            val alias = duplicatesMap!![key]
            val sourceRowNum = sourcesIndex[alias]!![key]!!
            val sourceRow = sources[alias]!!.getSheetAt(0).getRow(sourceRowNum)
            val targetKey = targetRow.getCell(targetIndexCol).stringCellValue
            val sourceKey = sourceRow.getCell(sourceIndexCol).stringCellValue
            //             cross-check control
            if (key == sourceKey && key == targetKey) {
                updateRow(targetRow, sourceRow, map)
            } else {
                println("mismatch in updating the keys! Duplicates contains: $key, targetKey: $targetKey, sourceKey: $sourceKey")
            }
            val data = mapOf(16 to "Dominiando", 17 to alias!!, 18 to alias + " SRL", 19 to key, 23 to key)
            suggestCellData(targetRow, data)
            markRow(targetRow, 25, markerForDuplicates, styleForDuplicates)
        }

    }

    /**
     * Fill in given row cells with provided strings. If a cell exists, it is skipped.

     * @param row  the cells of this row are to be updated
     * *
     * @param data map from the cell index (zero based) to the string content it should contain. If a cell exists,
     * *             its content is not modified.
     */
    private fun suggestCellData(row: Row, data: Map<Int, String>) {
        var cell: Cell?
        for (pos in data.keys) {
            cell = row.getCell(pos)
            if (cell == null) {
                cell = row.createCell(pos, Cell.CELL_TYPE_STRING)
                cell!!.setCellValue(data[pos])
            }
        }
    }

    /**
     * Adds a string cell at the end of the row which key is not present in any of the source files.
     */
    private fun updatesMissing() {
        val map = HashMap<Int, Int>()
        for (key in missingList!!) {
            val rowNum = targetIndex!![key]
            val row = target.getSheetAt(0).getRow(rowNum!!)
            updateRow(row, null, map)
            markRow(row, 25, markerForMissing, styleForMissing)
        }


    }

    private fun updateExtra() {
        for (key in extraMap!!.keys) {
            val alias = extraMap!![key]
            val sourceRowNum = sourcesIndex!![alias]!![key]
            val sourceRow = sources[alias]!!.getSheetAt(0).getRow(sourceRowNum!!)
            val totalRowNum = target.getSheetAt(0).lastRowNum
            val targetRow = target.getSheetAt(0).createRow(totalRowNum + 1)
            targetRow.createCell(targetIndexCol, Cell.CELL_TYPE_STRING).setCellValue(key)
            updateRow(targetRow, sourceRow, map)
            // set up by hand
            val data = mapOf(2 to "Confermato", 3 to "Confermato", 13 to "Sì", 16 to "Dominiando", 17 to alias!!, 18 to alias + " SRL", 19 to key, 23 to key)

            try {
                fillInRowCells(targetRow, data)
            } catch (e: Exception) {
                println("Error when adjusting an extraMap row for " + key + " from " + alias + ": " + e.message)
            }

            targetRow.getCell(5).setCellStyle(dateCellStyle)
            targetRow.getCell(6).setCellStyle(dateCellStyle)
            targetRow.getCell(7).setCellStyle(dateCellStyle)
            val cell = targetRow.createCell(22)
            cell.setCellValue(targetRow.getCell(7).dateCellValue)
            cell.setCellStyle(dateCellStyle)

            markRow(targetRow, 25, markerForExtra, styleForExtra)

        }
    }

    /**
     * Create cells in the row and fill them in with given strings.

     * @param row  the row whose cell are to be filled in
     * *
     * @param data map from cell numbers to string that the cell should contain.
     * *
     * @throws Exception if the row already contains at least one cell that should be filled in.
     */
    @Throws(Exception::class)
    private fun fillInRowCells(row: Row, data: Map<Int, String>) {
        for (index in data.keys) {
            var cell: Cell? = row.getCell(index)
            if (cell == null) {
                cell = row.createCell(index, Cell.CELL_TYPE_STRING)
            } else {
                throw Exception("Cell n. " + index + " already exists! It contains: " + cell.stringCellValue)
            }
            cell!!.setCellValue(data[index])
        }

    }

    /**
     * Updates targetRow with information from the sourceRow using given map as a correspondence between the row cells.

     * @param targetRow
     * *
     * @param sourceRow
     * *
     * @param map
     */
    private fun updateRow(targetRow: Row, sourceRow: Row?, map: Map<Int, Int>) {
        for (targetCellNum in map.keys) {
            val sourceCellNum = map[targetCellNum]
            val sourceCell = sourceRow!!.getCell(sourceCellNum!!)
            if (sourceCell == null) {
                println("source column $sourceCellNum is not present. Skipping it.")
                continue
            }
            val sourceCellType = sourceCell.cellType
            var targetCell: Cell? = targetRow.getCell(targetCellNum)

            if (targetCell != null && sourceCellType != targetCell.cellType) {
                println("cell type mismatch: source cell n. " + sourceCellNum + " of type " + sourceCell.cellType + " vs target cell n. " + targetCellNum
                        + " of type " + targetCell.cellType
                        + " for key " + targetRow.getCell(targetIndexCol).stringCellValue + ". Skipping it.")
                continue
            }
            if (targetCell == null) {
                targetCell = targetRow.createCell(targetCellNum, sourceCellType)
            }

            when (sourceCellType) {
                Cell.CELL_TYPE_BLANK -> println("source cell is blank")
                Cell.CELL_TYPE_BOOLEAN -> targetCell!!.setCellValue(sourceCell.booleanCellValue)
                Cell.CELL_TYPE_NUMERIC -> targetCell!!.setCellValue(sourceCell.numericCellValue)
                Cell.CELL_TYPE_STRING -> targetCell!!.setCellValue(sourceCell.stringCellValue)
                else -> println("Cell type $sourceCellType is not supported. Skipping the update of this cell.")
            }
        }


    }


    /**
     * Insert a marker at a cell with given number and apply cell styles.

     * @param targetRow
     * *
     * @param marker
     * *
     * @param style
     * *
     * @param cellNum   cell number (zero based) at which to insert the marker. -1 in order to insert at the end of the row.
     */
    private fun markRow(targetRow: Row, cellNum: Int, marker: String?, style: CellStyle?) {
        val pos = if (cellNum == -1) targetRow.lastCellNum + 1 else cellNum
        var cell: Cell? = targetRow.getCell(cellNum)
        if (cell == null) {
            cell = targetRow.createCell(pos, Cell.CELL_TYPE_STRING)
        }
        if (marker != null) {
            cell!!.setCellValue(marker)
        }
        if (style != null) {
            cell!!.cellStyle = style
        }
    }

}
