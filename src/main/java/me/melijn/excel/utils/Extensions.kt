package me.melijn.excel.utils

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import java.util.regex.Pattern
import kotlin.math.roundToLong

val Cell.stringCellValueBlank: String
    get() {
        return this.stringCellValue ?: ""
    }

val Cell.rawCellValue: String
    get() {
        return when (this.cellType) {
            CellType.BLANK -> ""
            CellType.BOOLEAN -> this.booleanCellValue.toString()
            CellType.STRING -> this.stringCellValue
            CellType.NUMERIC -> this.numericCellValue.fmt()
            CellType.FORMULA -> this.cellFormula
            else -> "not handled"
        }
    }

fun Cell.rawCellValue(sRow: Int? = null, nRow: Int? = null): String {
    val raw = this.rawCellValue
    return if (this.cellType == CellType.FORMULA && sRow != null && nRow != null) {
        raw.specialReplace(getColumnLetter(sRow), getColumnLetter(nRow))
    } else {
        raw
    }
}

private fun String.specialReplace(columnLetter: String, columnLetter1: String): String {
    val matcher = Pattern.compile("([A-Z]+)\\((.*)\\)").matcher(this)
    return if (matcher.find()) {
        val group = matcher.group(2)
        val newFormulaPart = group.replace(columnLetter, columnLetter1)
        "${matcher.group(1)}($newFormulaPart)"
    } else {
        this
    }
}



fun getColumnLetter(columnId: Int): String {
    // https://stackoverflow.com/a/182924/6160062
    var dividend = columnId + 1
    var columnLetters = ""
    var modulo = 0
    while (dividend > 0) {
        modulo = (dividend - 1) % 26
        columnLetters = (modulo + 65).toChar() + columnLetters
        dividend = (dividend - modulo) / 26
    }

    return columnLetters
}

fun Double.fmt(): String {
    return if (this == this.roundToLong().toDouble()) {
        String.format("%d", this.roundToLong())
    } else {
        String.format("%s", this)
    }
}