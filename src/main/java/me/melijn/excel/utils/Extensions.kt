package me.melijn.excel.utils

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
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
            else -> "not handled"
        }
    }

fun Double.fmt(): String {
    return if (this == this.roundToLong().toDouble()) {
        String.format("%d", this.roundToLong())
    } else {
        String.format("%s", this)
    }
}