package me.melijn.excel

import me.melijn.excel.internals.CellInfo
import me.melijn.excel.internals.FontOptions
import me.melijn.excel.utils.rawCellValue
import me.melijn.excel.utils.stringCellValueBlank
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
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
                    lRekeningCellName = lRekening.rawCellValue
                    docNrCellName = docNr.rawCellValue
                    referentiesCellName = referentie.rawCellValue
                    docDatumCellName = docDatum.rawCellValue
                    naamCellName = naam.rawCellValue
                    rekeningNummerCellName = rekeningNummer.rawCellValue
                    brutoBedragCellName = brutoBedrag.rawCellValue
                    valdotCellName = valdot.rawCellValue
                    btwCellName = btw.rawCellValue
                    nettoFactCellName = nettoFact.rawCellValue
                    kostenCellName = kosten.rawCellValue
                    basisRvCellName = basisRv.rawCellValue
                    percentRvCellName = percentRv.rawCellValue
                    teBetLevCellName = teBetLev.rawCellValue

                    map[row.rowNum] = arrayOf(
                        CellInfo(rekeningNummerCellName, rekeningNummer.cellType, FontOptions(22)),
                        CellInfo(teBetLevCellName, teBetLev.cellType, FontOptions(22)),
                        CellInfo(referentiesCellName, referentie.cellType, FontOptions(22)),
                        CellInfo(naamCellName, naam.cellType),
                        CellInfo(lRekeningCellName, lRekening.cellType),
                        CellInfo(docNrCellName, docNr.cellType),
                        CellInfo(docDatumCellName, docDatum.cellType),
                        CellInfo("", CellType.STRING),
                        CellInfo(brutoBedragCellName, brutoBedrag.cellType),
                        CellInfo(valdotCellName, valdot.cellType),
                        CellInfo(btwCellName, btw.cellType),
                        CellInfo(nettoFactCellName, nettoFact.cellType),
                        CellInfo(kostenCellName, kosten.cellType),
                        CellInfo(basisRvCellName, basisRv.cellType),
                        CellInfo(percentRvCellName, percentRv.cellType)
                    )
                }
                rekeningNummer.stringCellValueBlank.isBlank() -> {

                    //fat row
                    map[row.rowNum] = arrayOf(
                        CellInfo("", font = FontOptions(22)),
                        CellInfo(
                            teBetLev.rawCellValue(teBetLev.columnIndex, 1),
                            teBetLev.cellType, FontOptions(22, bold = true),
                            HorizontalAlignment.RIGHT, format = teBetLev.cellStyle.dataFormat
                        ),
                        CellInfo("", font = FontOptions(22)),
                        CellInfo(""),
                        CellInfo(lRekening.rawCellValue, lRekening.cellType, FontOptions(bold = true)),
                        CellInfo(""),
                        CellInfo(""),
                        CellInfo(""),
                        CellInfo(
                            brutoBedrag.rawCellValue(brutoBedrag.columnIndex, 8), brutoBedrag.cellType,
                            horizontalAlignment = HorizontalAlignment.RIGHT, format = brutoBedrag.cellStyle.dataFormat
                        ),
                        CellInfo(""),
                        CellInfo(""),
                        CellInfo(nettoFact.rawCellValue(nettoFact.columnIndex, 11), nettoFact.cellType, horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(""),
                        CellInfo(""),
                        CellInfo(percentRv.rawCellValue(nettoFact.columnIndex, 14), percentRv.cellType, horizontalAlignment = HorizontalAlignment.RIGHT)
                    )
                }
                else -> {

                    //normal row
                    val referentieRaw = referentie.rawCellValue
                    val referentieS = if (referentieRaw.length == 12 && referentieRaw.matches("\\d+".toRegex())) {
                        val firstThree = referentieRaw.substring(0, 3)
                        val secondFour = referentieRaw.substring(3, 8)
                        val thirdFour = referentieRaw.substring(8, 12)
                        "$firstThree/$secondFour/$thirdFour"
                    } else referentieRaw

                    map[row.rowNum] = arrayOf(
                        CellInfo(rekeningNummer.rawCellValue.insertCharEach(4, ' '), font = FontOptions(22)),
                        CellInfo(
                            teBetLev.rawCellValue, teBetLev.cellType, FontOptions(22),
                            HorizontalAlignment.RIGHT, format = teBetLev.cellStyle.dataFormat
                        ),
                        CellInfo(referentieS, font = FontOptions(22)),
                        CellInfo(naam.rawCellValue, naam.cellType),
                        CellInfo(lRekening.rawCellValue, lRekening.cellType, horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(docNr.rawCellValue, docNr.cellType),
                        CellInfo(docDatum.rawCellValue, horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(""),
                        CellInfo(
                            brutoBedrag.rawCellValue, brutoBedrag.cellType,
                            horizontalAlignment = HorizontalAlignment.RIGHT, format = brutoBedrag.cellStyle.dataFormat
                        ),
                        CellInfo(""),
                        CellInfo(btw.rawCellValue, btw.cellType, horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(nettoFact.rawCellValue, nettoFact.cellType, horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(kosten.rawCellValue, kosten.cellType, horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(basisRv.rawCellValue, basisRv.cellType, horizontalAlignment = HorizontalAlignment.RIGHT),
                        CellInfo(percentRv.rawCellValue, percentRv.cellType, horizontalAlignment = HorizontalAlignment.RIGHT)
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
                when (cellInfo.cellType) {
                    CellType.STRING -> cell.setCellValue(cellInfo.contentString())
                    CellType.NUMERIC -> cell.setCellValue(cellInfo.contentDouble())
                    CellType.BOOLEAN -> cell.setCellValue(cellInfo.contentBoolean())
                    CellType.FORMULA -> cell.cellFormula = cellInfo.contentFormula()
                    else -> {}
                }

                val font = doc.createFont()
                val newStyle = doc.createCellStyle()
                newStyle.alignment = cellInfo.horizontalAlignment
                newStyle.verticalAlignment = cellInfo.verticalAlignment
                cellInfo.format?.let { newStyle.dataFormat = it }
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