package com.excelsort.service

import com.excelsort.ESConstants
import com.excelsort.model.ColumnName
import com.excelsort.model.TargetDataEntry
import com.excelsort.utils.CellStyling
import lombok.AllArgsConstructor
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import java.io.File
import java.io.FileOutputStream
import java.io.IOException
import java.math.BigDecimal


@Service
@AllArgsConstructor
class WorkbookWriterService {

    private final val headerRowIndex = 1

    private final val excelDataLoaderService = ExcelDataLoaderService()

    fun write() {
        val workbooksMap: Map<String, Workbook> = excelDataLoaderService.loadExcelWorkbooks()

        for (entry in workbooksMap.entries) {
            val workbook = writeWorkbookContent(entry.value.getSheetAt(0))
            if (workbook != null) persistFile(entry.key, workbook)
        }
    }

    private fun writeWorkbookContent(sourceSheet: Sheet): Workbook? {
        val targetDataEntries = ArrayList<TargetDataEntry>()
        var rowNumber = headerRowIndex + 1

        while (true) {
            val currentRow = sourceSheet.getRow(rowNumber)
            if (currentRow == null || currentRow.getCell(1) == null) break

            val nazwaProduktu = currentRow.getCell(1).stringCellValue

            val iloscCell = currentRow.getCell(4)
            val ilosc =
                if (iloscCell !== null && iloscCell.cellType == CellType.NUMERIC) iloscCell.numericCellValue.toBigDecimal() else null

            val dostepnoscCell = currentRow.getCell(5)
            val dostepnosc =
                if (dostepnoscCell != null && dostepnoscCell.cellType == CellType.NUMERIC) dostepnoscCell.numericCellValue.toBigDecimal() else null

            val eanCell = currentRow.getCell(3)
            val ean =
                if (eanCell != null && eanCell.cellType == CellType.NUMERIC) BigDecimal(eanCell.numericCellValue) else null

            targetDataEntries.add(TargetDataEntry(nazwaProduktu, ilosc, dostepnosc, ean))

            rowNumber++
        }
        if (targetDataEntries.isNotEmpty()) {
            targetDataEntries.sortBy { it.nazwaProduktu.lowercase() }
            return writeTargetEntriesToANewWorkbook(targetDataEntries, sourceSheet.sheetName)
        }
        return null
    }

    private fun writeTargetEntriesToANewWorkbook(
        targetDataEntries: List<TargetDataEntry>,
        sheetName: String
    ): Workbook {
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet(sheetName)

        sheet.createRow(0)
        val headerRow = sheet.createRow(1)

        val regularCells = ArrayList<XSSFCell>()

        val cellHeaderNp =
            headerRow.createCell(0, CellType.STRING)
        cellHeaderNp.setCellValue(ColumnName.NAZWA_PRODUKTU.columnName)
        regularCells.add(cellHeaderNp)

        val cellHeaderI =
            headerRow.createCell(1, CellType.STRING)
        cellHeaderI.setCellValue(ColumnName.ILOSC.columnName)
        regularCells.add(cellHeaderI)

        val cellHeaderD =
            headerRow.createCell(2, CellType.STRING)
        cellHeaderD.setCellValue(ColumnName.DOSTEPNOSC.columnName)
        regularCells.add(cellHeaderD)

        val cellHeaderEan =
            headerRow.createCell(3, CellType.STRING)
        cellHeaderEan.setCellValue(ColumnName.EAN.columnName)
        regularCells.add(cellHeaderEan)

        val nazwaProduktuCells = ArrayList<XSSFCell>()
        for ((index, entry) in targetDataEntries.withIndex()) {
            var cellNp: XSSFCell
            var cellI: XSSFCell
            var cellD: XSSFCell
            var cellEan: XSSFCell

            val row = sheet.createRow(index + 1 + headerRowIndex)

            cellNp =
                row.createCell(0, CellType.STRING)
            cellNp.setCellValue(entry.nazwaProduktu)
            nazwaProduktuCells.add(cellNp)

            if (entry.ilosc != null) {
                cellI = row.createCell(1, CellType.NUMERIC)
                cellI.setCellValue((entry.ilosc).toDouble())
            } else {
                cellI = row.createCell(1, CellType.BLANK)
            }
            regularCells.add(cellI)

            if (entry.dostepnosc != null) {
                cellD = row.createCell(2, CellType.NUMERIC)
                cellD.setCellValue((entry.dostepnosc).toDouble())
            } else {
                cellD = row.createCell(2, CellType.BLANK)
            }
            regularCells.add(cellD)

            if (entry.ean != null) {
                cellEan = row.createCell(3, CellType.STRING)
                cellEan.setCellValue(entry.ean.toDouble())
            } else {
                cellEan = row.createCell(3, CellType.BLANK)
            }
            regularCells.add(cellEan)
        }

        val regularStyle = CellStyling.getRegularStyle(workbook)
        regularCells.forEach { it.cellStyle = regularStyle }

        val nazwaProduktuStyle = CellStyling.getNazwaProduktuStyle(workbook)
        nazwaProduktuCells.forEach { it.cellStyle = nazwaProduktuStyle }

        sheet.setColumnWidth(0, 11_052); // Approximately 9.0 cm
        sheet.setColumnWidth(1, 1_842);  // Approximately 1.5 cm
        sheet.setColumnWidth(2, 3_070);  // Approximately 2.5 cm
        sheet.setColumnWidth(3, 4_298);  // Approximately 3.5 cm

        return workbook
    }

    private fun persistFile(fileName: String, workbook: Workbook) {
        val targetDir = File(ESConstants.ROOT_DIR, ESConstants.TARGET_FOLDER_NAME)
        val fullFileName = "$fileName - poprawione.xlsx"

        if (!targetDir.exists()) {
            targetDir.mkdir()
        }
        val targetFile = File(targetDir, fullFileName)

        try {
            FileOutputStream(targetFile).use { fileOut ->
                workbook.write(fileOut)
            }
            println("Workbook successfully persisted to: ${targetFile.absolutePath}")
        } catch (e: IOException) {
            System.err.println("Error persisting the workbook: ${e.message}")
            e.printStackTrace()
        }
    }

    private fun correctEanValue(ean: BigDecimal): Double {
        return ean.multiply(BigDecimal("10")).divide(BigDecimal("10")).toDouble()
    }


}