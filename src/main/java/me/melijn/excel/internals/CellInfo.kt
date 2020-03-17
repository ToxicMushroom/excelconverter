package me.melijn.excel.internals

import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment

data class CellInfo(
    val content: Any,
    val font: FontOptions,
    val horizontalAlignment: HorizontalAlignment = HorizontalAlignment.LEFT,
    val verticalAlignment: VerticalAlignment = VerticalAlignment.BOTTOM
)