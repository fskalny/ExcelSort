package com.example.excelsort.service

import com.example.excelsort.ESConstants
import com.example.excelsort.model.ColumnName
import com.example.excelsort.model.TargetDataEntry
import lombok.AllArgsConstructor
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import java.io.File
import java.io.FileOutputStream
import java.io.IOException


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
            val ilosc = if (iloscCell.cellType == CellType.NUMERIC) iloscCell.numericCellValue.toBigDecimal() else null

            val dostepnoscCell = currentRow.getCell(5)
            val dostepnosc =
                if (dostepnoscCell.cellType == CellType.NUMERIC) dostepnoscCell.numericCellValue.toBigDecimal() else null

            val eanCell = currentRow.getCell(3)
            val ean = if (dostepnoscCell.cellType == CellType.NUMERIC) eanCell.numericCellValue.toBigDecimal() else null

            targetDataEntries.add(TargetDataEntry(nazwaProduktu, ilosc, dostepnosc, ean))

            rowNumber++
        }
        if (targetDataEntries.isNotEmpty()) {
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

        headerRow.createCell(0, CellType.STRING).setCellValue(ColumnName.NAZWA_PRODUKTU.columnName)
        headerRow.createCell(1, CellType.STRING).setCellValue(ColumnName.ILOSC.columnName)
        headerRow.createCell(2, CellType.STRING).setCellValue(ColumnName.DOSTEPNOSC.columnName)
        headerRow.createCell(3, CellType.STRING).setCellValue(ColumnName.EAN.columnName)

        for ((index, entry) in targetDataEntries.withIndex()) {
            val row = sheet.createRow(index + 1 + headerRowIndex)
            row.createCell(0, CellType.STRING).setCellValue(entry.nazwaProduktu)

            if (entry.ilosc != null) row.createCell(1, CellType.NUMERIC).setCellValue((entry.ilosc).toDouble())
            if (entry.dostepnosc != null) row.createCell(2, CellType.NUMERIC)
                .setCellValue((entry.dostepnosc).toDouble())
            if (entry.ean != null) row.createCell(3, CellType.NUMERIC).setCellValue((entry.ean).toDouble())
        }

        for (idx in 0..3) {
            sheet.autoSizeColumn(idx)
        }

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


}