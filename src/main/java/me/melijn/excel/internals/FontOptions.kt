package me.melijn.excel.internals

import org.apache.poi.ss.usermodel.FontUnderline

data class FontOptions(
    val size: Short = 11,
    val family: String = "Calibri",
    val bold: Boolean = false,
    val italic: Boolean = false,
    val underline: FontUnderline = FontUnderline.NONE,
    val strikeout: Boolean = false
)