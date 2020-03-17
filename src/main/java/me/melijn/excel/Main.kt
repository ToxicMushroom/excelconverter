package me.melijn.excel

import me.melijn.excel.internals.CellInfo
import me.melijn.excel.internals.FontOptions
import me.melijn.excel.utils.rawCellValue
import me.melijn.excel.utils.stringCellValueBlank
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.util.*


class Main(args: Array<String>) {

    private val lRekeningColumnIndex = 1
    private val docNrColumnIndex = 2
    private val referentiesColumnIndex = 3
    private val docDatumColumnIndex = 4
    private val naamColumnIndex = 5
    private val rekeningNrColumnIndex = 6
    private val brutoBedragColumnIndex = 7
    private val valdotColumnIndex = 8
    private val btwColumnIndex = 9
    private val nettoFactColumnIndex = 10
    private val kostenColumnIndex = 11
    private val basisRvColumnIndex = 12
    private val percentRvColumnIndex = 13
    private val teBetLevColumnIndex = 14

    private lateinit var lRekeningCellName: String
    private lateinit var docNrCellName: String
    private lateinit var referentiesCellName: String
    private lateinit var docDatumCellName: String
    private lateinit var naamCellName: String
    private lateinit var rekeningNummerCellName: String
    private lateinit var brutoBedragCellName: String
    private lateinit var valdotCellName: String
    private lateinit var btwCellName: String
    private lateinit var nettoFactCellName: String
    private lateinit var kostenCellName: String
    private lateinit var basisRvCellName: String
    private lateinit var percentRvCellName: String
    private lateinit var teBetLevCellName: String
    private val missingCellPolicy = Row.MissingCellPolicy.CREATE_NULL_AS_BLANK

    init {
        // https://poi.apache.org/components/spreadsheet/quick-guide.html | basically everything you need
        val path = if (args.isEmpty()) {
            "base.xlsx"
        } else {
            args[0]
        }
        val file = File(path)
        val doc = XSSFWorkbook(file)
        val oldSheet = doc.getSheetAt(0)
        val newSheet = doc.createSheet("result")

        val map: MutableMap<Int, Array<CellInfo>> = mutableMapOf()
        for (row in oldSheet.rowIterator()) {

            val lRekening: Cell = row.getCell(lRekeningColumnIndex, missingCellPolicy)
            val docNr: Cell = row.getCell(docNrColumnIndex, missingCellPolicy)
            val referentie: Cell = row.getCell(referentiesColumnIndex, missingCellPolicy)
            val docDatum: Cell = row.getCell(docDatumColumnIndex, missingCellPolicy)
            val naam: Cell = row.getCell(naamColumnIndex, missingCellPolicy)
            val rekeningNummer: Cell = row.getCell(rekeningNrColumnIndex, missingCellPolicy)
            val brutoBedrag: Cell = row.getCell(brutoBedragColumnIndex, missingCellPolicy)
            val valdot: Cell = row.getCell(valdotColumnIndex, missingCellPolicy)
            val btw: Cell = row.getCell(btwColumnIndex, missingCellPolicy)
            val nettoFact: Cell = row.getCell(nettoFactColumnIndex, missingCellPolicy)
            val kosten: Cell = row.getCell(kostenColumnIndex, missingCellPolicy)
            val basisRv: Cell = row.getCell(basisRvColumnIndex, missingCellPolicy)
            val percentRv: Cell = row.getCell(percentRvColumnIndex, missingCellPolicy)
            val teBetLev: Cell = row.getCell(teBetLevColumnIndex, missingCellPolicy)

            when {
                row.rowNum == 0 -> {
                    //Title row
                    lRekeningCellName = lRekening.stringCellValue
                    docNrCellName = docNr.stringCellValue
                    referentiesCellName = referentie.stringCellValue
                    docDatumCellName = docDatum.stringCellValue
                    naamCellName = naam.stringCellValue
                    rekeningNummerCellName = rekeningNummer.stringCellValue
                    brutoBedragCellName = brutoBedrag.stringCellValue
                    valdotCellName = valdot.stringCellValue
                    btwCellName = btw.stringCellValue
                    nettoFactCellName = nettoFact.stringCellValue
                    kostenCellName = kosten.stringCellValue
                    basisRvCellName = basisRv.stringCellValue
                    percentRvCellName = percentRv.stringCellValue
                    teBetLevCellName = teBetLev.stringCellValue

                    map[row.rowNum] = arrayOf(
                        CellInfo(rekeningNummerCellName, FontOptions(22)),
                        CellInfo(teBetLevCellName, FontOptions(22)),
                        CellInfo(referentiesCellName, FontOptions(22)),
                        CellInfo(naamCellName, FontOptions()),
                        CellInfo(lRekeningCellName, FontOptions()),
                        CellInfo(docNrCellName, FontOptions()),
                        CellInfo(docDatumCellName, FontOptions()),
                        CellInfo("", FontOptions()),
                        CellInfo(brutoBedragCellName, FontOptions()),
                        CellInfo(valdotCellName, FontOptions()),
                        CellInfo(btwCellName, FontOptions()),
                        CellInfo(nettoFactCellName, FontOptions()),
                        CellInfo(kostenCellName, FontOptions()),
                        CellInfo(basisRvCellName, FontOptions()),
                        CellInfo(percentRvCellName, FontOptions())
                    )
                }
                rekeningNummer.stringCellValueBlank.isBlank() -> {

                    //fat row
                    map[row.rowNum] = arrayOf(
                        CellInfo("", FontOptions(22)),
                        CellInfo(teBetLev.numericCellValue, FontOptions(22, bold = true), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo("", FontOptions(22)),
                        CellInfo("", FontOptions()),
                        CellInfo(lRekening.stringCellValue, FontOptions(bold = true)),
                        CellInfo("", FontOptions()),
                        CellInfo("", FontOptions()),
                        CellInfo("", FontOptions()),
                        CellInfo(brutoBedrag.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo("", FontOptions()),
                        CellInfo("", FontOptions()),
                        CellInfo(nettoFact.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo("", FontOptions()),
                        CellInfo("", FontOptions()),
                        CellInfo(percentRv.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT)
                    )
                }
                else -> {

                    //normal row
                    val referentieRaw = referentie.stringCellValue
                    val referentieS = if (referentieRaw.length == 12 && referentieRaw.matches("\\d+".toRegex())) {
                        val firstThree = referentieRaw.substring(0, 3)
                        val secondFour = referentieRaw.substring(3, 8)
                        val thirdFour = referentieRaw.substring(8, 12)
                        "$firstThree/$secondFour/$thirdFour"
                    } else referentieRaw

                    map[row.rowNum] = arrayOf(
                        CellInfo(rekeningNummer.stringCellValue.insertCharEach(4, ' '), FontOptions(22)),
                        CellInfo(teBetLev.numericCellValue, FontOptions(22), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(referentieS, FontOptions(22)),
                        CellInfo(naam.stringCellValue, FontOptions()),
                        CellInfo(lRekening.rawCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(docNr.stringCellValue, FontOptions()),
                        CellInfo(docDatum.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo("", FontOptions()),
                        CellInfo(brutoBedrag.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo("", FontOptions()),
                        CellInfo(btw.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(nettoFact.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(kosten.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(basisRv.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(percentRv.numericCellValue, FontOptions(), horizontalAlignment = HorizontalAlignment.RIGHT)
                    )
                }
            }
        }

        val keys = map.keys

        for ((rowIndex, key) in keys.withIndex()) {
            val row = newSheet.createRow(rowIndex)
            val cellArray = map[key] ?: continue

            for ((cellIndex, cellInfo) in cellArray.withIndex()) {
                val cell = row.createCell(cellIndex)
                when (cellInfo.content) {
                    is String -> cell.setCellValue(cellInfo.content)
                    is Double -> cell.setCellValue(cellInfo.content)
                    is Boolean -> cell.setCellValue(cellInfo.content)
                    is Date -> cell.setCellValue(cellInfo.content)
                    is Calendar -> cell.setCellValue(cellInfo.content)
                    else -> cell.setCellValue(cellInfo.content.toString())
                }

                val font = doc.createFont()
                val newStyle = doc.createCellStyle()
                newStyle.alignment = cellInfo.horizontalAlignment
                newStyle.verticalAlignment = cellInfo.verticalAlignment
                font.bold = cellInfo.font.bold
                font.fontHeightInPoints = cellInfo.font.size
                newStyle.setFont(font)


                cell.cellStyle = newStyle
                newSheet.autoSizeColumn(cell.columnIndex)
            }
        }


        val fileOut = if (file.parentFile.isDirectory) {
            File("${file.parentFile.absolutePath}/resultaat.xlsx")
        } else {
            File("resultaat.xlsx")
        }
        FileOutputStream(fileOut).use {
            doc.write(it)
        }
    }
}

private fun String.insertCharEach(amount: Int, c: Char): String {
    var newS = ""

    var counter = 0
    for (char in this.toCharArray()) {
        newS += char
        if (++counter == amount) {
            newS += c
            counter = 0
        }
    }
    return newS
}

fun main(args: Array<String>) {
    Main(args)
}