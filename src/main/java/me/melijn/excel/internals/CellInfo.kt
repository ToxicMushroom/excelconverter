package me.melijn.excel.internals

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment

data class CellInfo(
    val content: Any,
    val cellType: CellType = CellType.STRING,
    val font: FontOptions = FontOptions(),
    val horizontalAlignment: HorizontalAlignment = HorizontalAlignment.LEFT,
    val verticalAlignment: VerticalAlignment = VerticalAlignment.BOTTOM,
    val format: Short? = null
) {
    fun contentString(): String = content.toString()
    fun contentDouble(): Double = content.toString().toDouble()
    fun contentBoolean(): Boolean = content.toString().toBoolean()
    fun contentFormula(): String = "$content"
}